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
    python tools/dokaz.py amb-pisac --izbor    # pusta samo SVOJ test, ne celu suite
    python tools/dokaz.py --izolacija          # koji ciljani testovi ne prolaze SAMI

CETIRI STVARI KOJE OVAJ ALAT TVRDI, a ne samo gleda:

1. BAZA MORA BITI ZELENA PRE PRVE MUTACIJE. Bez toga alat ne dokazuje "mutacija
   je izazvala crveno" nego samo "posle mutacije postoji crveno" -- a to nisu
   iste tvrdnje. Test koji vec pada iz nekog treceg razloga proglasio bi svaku
   sabotazu nad sobom dokazanom, ukljucujuci onu koja ne radi nista.
   "Pre" je pre OCENE, ne nuzno pre pokretanja: baza i mutacija idu nad istim
   snimkom, pa uz radnike (R-4) mutacija cele suite sme da krene dok baza
   unapred jos traje -- ali se priznaje tek kad je baza poznata i zelena.

2. RADNI IZVOR SE NE DIRA (D-6). Na pocetku se `src-vba/` preslika u SNIMAK, a
   mutacije idu u KOPIJU snimka u privremenom folderu (`sabotaza.py --src`,
   `run_vba.py --src`). Radno stablo tokom dokaza niko ne pise: operater sme da
   radi, hook ostaje cist, a prekinut dokaz ne ostavlja pokvaren izvor. Posle
   svakog prolaza se kopija vraca iz snimka i poredi se njen potpis -- ako se ne
   poklopi ni posle obnove cele kopije, dokaz STAJE, jer bi sve mereno posle
   toga islo nad pokvarenim kodom. Snimak sa vec primenjenom sabotazom (ostatak
   rucnog dokaza) se odbija pre prve baze. Dokaz meri SNIMAK: ako se radno
   stablo menjalo tokom dokaza, verdikt imenuje snimak.

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

5. IZBOR (`--izbor`) PUSTA SAMO CILJANE TESTOVE, ALI SA SVOJOM BAZOM. Prolaz
   suite-a (RunAllTests 215,8 s pri 336 testova, 05.11.2026; BFP preko 300 s)
   zamenjuje prolaz samo testova clanova (`run_vba --test`). Tacka 1 se tada
   tvrdi nad ISTIM izborom: pre
   mutacije se isti spisak testova pusti nad cistim izvorom i mora biti zelen.
   Bez toga bi test koji SAM pada (zavisi od stanja koje mu ostavlja raniji
   test -- redosled u RunAllTests je nosec) "dokazao" svaku sabotazu nad sobom.
   Takav clan se ne gubi: grupa se rasturi na pojedinacna merenja, a test koji
   ni sam ne prolazi meri se CELOM suite-om, kao bez `--izbor`. Isto vazi za
   sabotazu koju izbor NE obori: test pusten sam moze da prodje mimo ciljane
   tvrdnje, pa "NIJE DOKAZANO" daje tek cela suite. Baza se pamti po izboru, pa
   vise sabotaza nad istim testom placa jednu.

   Izbor tvrdi MANJE od pune suite: dokazuje da je test osetljiv na mutaciju
   kad ide sam, ne i u kanonskom redosledu, gde raniji test moze da ostavi
   stanje koje mutaciju zaklanja. Verdikt se zato zove `DOKAZANO (izbor)`
   (uz grupe `DOKAZANO (grupno, izbor)`); golo `DOKAZANO` ostaje samo za
   pojedinacan dokaz u celoj suiti.

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
import concurrent.futures
import contextlib
import hashlib
import io
import importlib.util
import inspect
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time

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

# Radni folder tekuceg dokaza (snimak, kopija, zamrznut katalog). Pamti se samo
# zbog poruke pri prekidu -- izvor koji treba vratiti vise ne postoji.
_RADNI = None

# --- R-4: radnici ----------------------------------------------------------
#
# Svaki radnik ima SVOJU kopiju izvora i SVOJ izlaz run_vba-a (izvestaj,
# istorija, excel.pid). Posao (grupa ili jedna sabotaza) uzima slobodnog
# radnika, primeni-pusti-obnovi nad njegovom kopijom, i vraca ga. Rezultati se
# ispisuju i knjize U REDOSLEDU KATALOGA, kao serijski -- pa je --radnika 1
# isto sto i dokaz bez radnika, a N > 1 se sa njim poredi diff-om.
#
# Paralelno smeju samo suite za koje je to provereno (svaka u svom Excel-u,
# nad svojom temp kopijom sveske). Suite koje idu na mrezu ili dele stanje van
# sveske ostaju u serijskoj traci (--radnika 1).
PARALELNE_SUITE = {SUITE_ALL, SUITE_BANKA, SUITE_BFP}
_STOP = threading.Event()          # STOP ili Ctrl+C: nijedan nov posao ne krece
_START = {"lock": threading.Lock(), "poslednji": 0.0, "razmak": 0.0,
          "ponovljeno": 0}
_BAZE_LOCK = threading.Lock()
_RV_MODUL = {}


def _modul_sabotaza():
    put = os.path.join(ROOT, "tools", "sabotaza.py")
    spec = importlib.util.spec_from_file_location("_sab_za_dokaz", put)
    modul = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modul)
    return modul


def _run_vba_modul():
    """tools/run_vba.py kao modul (jednom po procesu)."""
    if "m" not in _RV_MODUL:
        put = os.path.join(ROOT, "tools", "run_vba.py")
        spec = importlib.util.spec_from_file_location("_run_vba_za_dokaz", put)
        modul = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(modul)
        _RV_MODUL["m"] = modul
    return _RV_MODUL["m"]


def _suite_izbor() -> set:
    """Suite koje primaju `run_vba --test` -- iz SUITES (jedini izvor istine).

    Cita se iz kataloga, ne prepisuje: druga lista bi bila druga stvar koja
    moze da se razidje sa ulaznim tackama u VBA (v. vba_gate, nalaz IZBOR).
    """
    modul = _run_vba_modul()
    return {ime for ime, meta in modul.SUITES.items() if meta.get("izbor")}


def _run_suite(suite: str, testovi=None, bez_markera: bool = False,
               src: str = None, izlaz: str = None) -> tuple:
    """Argumenti za `run_vba.py`: cela suite, ili samo `testovi` iz nje.

    `bez_markera` je za prolaz nad MUTIRANIM izvorom (KI-036): zelen prolaz u
    kome sabotaza ne obori nista inace upise marker zelenog nad izvorom koji
    nije cist. `src` je kopija izvora koju dokaz meri (D-6); ide NA KRAJ, pa
    raspored ostalih argumenata ostaje isti.
    """
    a = (sys.executable, "tools/run_vba.py", "--suite", suite)
    for test in testovi or ():
        a += ("--test", test)
    if bez_markera:
        a += ("--bez-markera",)
    if izlaz:
        a += ("--izlaz", izlaz)
    if src:
        a += ("--src", src)
    return a


def _fatal(izlaz: str) -> str:
    """Tekst reda FATAL iz izlaza run_vba, ili "" -- Excel je ubijen ili pukao.

    Bez ovoga je ubijen prolaz izgledao kao "NE OBARA NISTA" (nema FAIL redova)
    ili kao "nema oznacenog reda TESTS", a razlog (faza, dijalog) se gubio.
    """
    m = re.search(r"^FATAL\s+(.*)$", izlaz, re.M)
    return m.group(1).strip() if m else ""


def _otisak(put: str = None) -> str:
    """Potpis celog izvora u `put` (podrazumevano src-vba) -- imena i bajtovi."""
    put = put or SRC_VBA
    h = hashlib.sha256()
    for ime in sorted(os.listdir(put)):
        h.update(ime.encode())
        with open(os.path.join(put, ime), "rb") as fh:
            h.update(fh.read())
    return h.hexdigest()[:16]


# --- D-6: dokaz radi nad kopijom, radno stablo samo cita ----------------------


def _u_stablu(put: str, koren: str) -> bool:
    """Da li je `put` bas `koren` ili nesto ispod njega (realpath, normcase)."""
    p = os.path.normcase(os.path.realpath(put))
    k = os.path.normcase(os.path.realpath(koren))
    return p == k or p.startswith(k.rstrip(os.sep) + os.sep)


def _van_radnog_stabla(put: str, ctx: dict = None) -> bool:
    """STRAZA: cilj mutacije nije src-vba, ni snimak, ni nesto ispod njih.

    Snimak je zajednicki izvor iz koga se kopija vraca; mutacija u njemu bi se
    prenela u svaki sledeci prolaz. Zato je zabranjen isto kao radno stablo.
    """
    zabranjeni = [SRC_VBA] + ([ctx["snimak"]] if ctx else [])
    return not any(_u_stablu(put, z) for z in zabranjeni)


def _pripremi_radni(sab, koren_src: str = None, radnika: int = 1) -> tuple:
    """(ctx, problem): snimak izvora, zamrznut katalog i kopija za mutacije.

    Raspored je namerno takav da zamrznut `sabotaza.py` (radni/alat/) bez
    `--src` gadja `radni/src-vba`, koji NE POSTOJI: poziv koji zaboravi `--src`
    padne, umesto da tiho pokvari snimak ili radno stablo.
    """
    koren_src = koren_src or SRC_VBA
    radni = tempfile.mkdtemp(prefix="dokaz_")
    ctx = {"radni": radni,
           "snimak": os.path.join(radni, "snimak", "src-vba"),
           "kopija": os.path.join(radni, "r1", "src-vba"),
           "izlaz": os.path.join(radni, "r1", "izlaz"),
           "alat": os.path.join(radni, "alat", "sabotaza.py"),
           "id": 1}
    # Radnik K ima svoju kopiju i svoj izlaz; snimak i zamrznut katalog dele.
    ctx["bazen"] = [ctx] + [dict(ctx, id=k,
                                 kopija=os.path.join(radni, "r%d" % k, "src-vba"),
                                 izlaz=os.path.join(radni, "r%d" % k, "izlaz"))
                            for k in range(2, max(1, radnika) + 1)]
    for c in ctx["bazen"][1:]:
        c.pop("bazen", None)
    # Citanje ili kopiranje koje pukne je STOP sa razlogom, ne traceback; svaki
    # drugi izlaz (i Ctrl+C usred kopiranja) brise polu-napravljen folder.
    try:
        return ctx, _napuni_radni(ctx, sab, koren_src)
    except (OSError, UnicodeDecodeError) as exc:
        return ctx, "priprema snimka nije uspela -- %s: %s" % (type(exc).__name__, exc)
    except BaseException:
        shutil.rmtree(radni, ignore_errors=True)
        raise


def _napuni_radni(ctx: dict, sab, koren_src: str) -> str:
    """Telo pripreme; vraca problem ili ""."""
    for _pokusaj in (1, 2):
        shutil.rmtree(ctx["snimak"], ignore_errors=True)
        koren_pre = _otisak(koren_src)
        shutil.copytree(koren_src, ctx["snimak"])
        if _otisak(ctx["snimak"]) == koren_pre == _otisak(koren_src):
            break
    else:
        return ("src-vba se menjao dok se pravio snimak -- pusti dokaz ponovo "
                "kad izmena stane")
    zatecene = sab.zatecene(ctx["snimak"])
    if zatecene:
        return ("izvor je ZATECEN SABOTIRAN (%s%s) -- svaka mutacija bi se merila "
                "nad tom. Vrati: python tools/sabotaza.py --vrati"
                % (", ".join(zatecene[:5]), ", ..." if len(zatecene) > 5 else ""))
    os.makedirs(os.path.dirname(ctx["alat"]))
    shutil.copy2(os.path.join(ROOT, "tools", "sabotaza.py"), ctx["alat"])
    for c in ctx["bazen"]:
        shutil.copytree(ctx["snimak"], c["kopija"])
        os.makedirs(c["izlaz"])
    return _proveri_bazen(ctx["bazen"], ctx)


def _proveri_bazen(bazen: list, ctx: dict) -> str:
    """Radnici ne smeju da dele kopiju ni izlaz, niti da pisu u src-vba/snimak."""
    kopije = [os.path.normcase(os.path.realpath(c["kopija"])) for c in bazen]
    izlazi = [os.path.normcase(os.path.realpath(c["izlaz"])) for c in bazen]
    if len(set(kopije)) != len(kopije) or len(set(izlazi)) != len(izlazi):
        return "STRAZA: dva radnika dele kopiju ili izlaz"
    if any(not _van_radnog_stabla(c["kopija"], ctx) for c in bazen):
        return "STRAZA: kopija za mutacije je u radnom stablu ili u snimku"
    if any(_u_stablu(c["izlaz"], SRC_VBA) or _u_stablu(c["izlaz"], ctx["snimak"])
           for c in bazen):
        return "STRAZA: izlaz radnika je u radnom stablu ili u snimku"
    return ""


def _koren_menjan(snimak: str, tokom, posle: str) -> bool:
    """Da li se src-vba menjao OD SNIMKA: tokom nekog prolaza ili do kraja.

    Poredi se sa potpisom snimka, ne sa uzorkom uzetim kasnije: priprema je
    tvrdila da su snimak i src-vba isti u trenutku kopiranja, pa izmena tokom
    baza (minuti) mora da se racuna -- uzorak posle baza bi je vec sadrzao i
    proglasio stablo "nepromenjenim".
    """
    return any(k and k != snimak for k in tokom) or posle != snimak


def _vrati_kopiju(ctx: dict, fajlovi, pre: str) -> str:
    """Obnovi kopiju iz snimka i vrati njen potpis (jednak `pre` = uspeh).

    Prvo samo `fajlovi`; ako potpis ne legne (zakljucan fajl, mutacija van
    svog fajla) -- cela kopija, pa jos jednom.
    """
    sada = ""
    try:
        _obnovi(ctx, fajlovi)
        sada = _otisak(ctx["kopija"])
    except OSError:
        sada = ""
    if sada != pre:
        try:
            _obnovi(ctx, None)
            sada = _otisak(ctx["kopija"])
        except OSError as exc:
            sada = "nije obnovljena: %s" % exc
    return sada


def _obnovi(ctx: dict, fajlovi=None) -> None:
    """Kopija se vraca iz snimka: samo `fajlovi`, ili cela (None)."""
    if fajlovi is None:
        shutil.rmtree(ctx["kopija"], ignore_errors=True)
        shutil.copytree(ctx["snimak"], ctx["kopija"])
        return
    for f in sorted(fajlovi):
        shutil.copy2(os.path.join(ctx["snimak"], f), os.path.join(ctx["kopija"], f))


# Mora biti VECE od tvrdog prekida run_vba-a (--timeout, 1800 s): inace proces
# run_vba ubije ovaj poziv pre nego sto run_vba sam prijavi FATAL, pa se umesto
# razloga vidi samo "timeout" (KI-096).
def _pusti(*a, timeout=2400):
    return subprocess.run(a, cwd=ROOT, capture_output=True, text=True,
                          timeout=timeout)


def _kapija_starta(sat=time.monotonic, spavaj=time.sleep) -> None:
    """Najmanje `razmak` sekundi izmedju dva run_vba poziva (svi radnici).

    run_vba vec serijalizuje sam start Excela (start brava); razmak dodatno
    rasporedjuje pripremu bez Excela (kopija sveske, otisci), koja bi inace
    kod pet radnika krenula u istoj sekundi."""
    with _START["lock"]:
        cekaj = _START["razmak"] - (sat() - _START["poslednji"])
        if cekaj > 0:
            spavaj(cekaj)
        _START["poslednji"] = sat()


def _pusti_suite(*a, izlaz_dir: str = None, prekini=None):
    """run_vba poziv: kapija starta, i JEDAN ponovljen pokusaj kad Excel ni ne
    krene (FATAL bez ijednog reda SUITE -- nista nije mereno).

    Pad POSLE starta suite-a se ne ponavlja: to je ishod (npr. sabotaza koja
    obori compile), ne infrastruktura. Posle STOP-a/Ctrl+C se ne ponavlja nista,
    a ni posle `prekini` (otkazana baza unapred: njen Excel je ubijen namerno,
    pa FATAL nije znak da treba probati ponovo).
    Kad dokaz sam prekine poziv (timeout), ubija se i Excel tog poziva -- ali
    samo ako se PID i vreme nastanka iz excel.pid poklope."""
    for pokusaj in (1, 2):
        _kapija_starta()
        if _STOP.is_set() or (prekini is not None and prekini.is_set()):
            # Posao je primljen pre STOP-a, ali run jos nije krenuo: ne krece.
            return subprocess.CompletedProcess(a, 1, "", "PREKID: run_vba nije "
                                               "pokrenut (STOP ili Ctrl+C)\n")
        try:
            r = _pusti(*a)
        except subprocess.TimeoutExpired:
            _ubij_zaostali(izlaz_dir)
            raise
        tekst = (r.stdout or "") + (r.stderr or "")
        if (pokusaj == 1 and _fatal(tekst) and not _STOP.is_set()
                and not (prekini is not None and prekini.is_set())
                and not re.search(r"^SUITE\s", tekst, re.M)):
            with _START["lock"]:
                _START["ponovljeno"] += 1
            continue
        return r


def _ubij_zaostali(izlaz_dir: str, otvori=None, vreme=None, ubij=None) -> bool:
    """Ubij Excel koji je run_vba ostavio, SAMO ako se PID i vreme nastanka iz
    excel.pid poklope sa zivim procesom. True = ubijen.

    Excel pokrenut preko DCOM-a nije dete python-a: kad dokaz ubije run_vba
    (timeout), Excel ostaje. Isti PID kasnije moze biti tudji -- zato vreme."""
    if not izlaz_dir:
        return False
    put = os.path.join(izlaz_dir, "excel.pid")
    try:
        with open(put, encoding="ascii") as fh:
            zapis = json.load(fh)
    except (OSError, ValueError):
        return False
    if otvori is None:
        if os.name != "nt":
            return False
        rv = _run_vba_modul()
        otvori, vreme = rv._otvori_proces, rv.vreme_nastanka
        ubij = lambda h: rv._terminate_handle(h, {})   # noqa: E731
    h = otvori(int(zapis.get("pid") or 0))
    if h is None or not zapis.get("kreiran") or vreme(h) != zapis["kreiran"]:
        return False
    ubij(h)
    try:
        os.remove(put)
    except OSError:
        pass
    return True


def _baza_izbora(baze: dict, kljuc, suite: str, testovi, ctx: dict) -> tuple:
    """Baza izbora, jednom po kljucu -- i kad je trazi vise radnika odjednom.

    Prvi koji je trazi je meri (nad SVOJOM kopijom, jednakom snimku); ostali
    cekaju njegov rezultat. Bez ovoga bi dva radnika platila istu bazu dvaput."""
    with _BAZE_LOCK:
        v = baze.get(kljuc)
        if v is None:
            dogadjaj = threading.Event()
            baze[kljuc] = dogadjaj
    if v is None:
        try:
            rez = _baza_zelena(suite, testovi, src=ctx["kopija"],
                               izlaz=ctx.get("izlaz"))
        except BaseException:
            with _BAZE_LOCK:
                baze.pop(kljuc, None)
            dogadjaj.set()
            raise
        with _BAZE_LOCK:
            baze[kljuc] = rez
        dogadjaj.set()
        return rez
    if isinstance(v, threading.Event):
        v.wait()
        with _BAZE_LOCK:
            gotovo = baze.get(kljuc)
        if isinstance(gotovo, tuple):
            return gotovo
        return _baza_izbora(baze, kljuc, suite, testovi, ctx)
    return v


def _izvrsi(poslovi: list, bazen: list):
    """Pusti poslove (fn(ctx) -> rezultat) nad bazenom radnika. GENERATOR:
    rezultati stizu U REDOSLEDU POSLOVA, cim je sledeci po redu gotov.

    Radnik (ctx) je u jednom trenutku kod najvise jednog posla. Kad je _STOP
    postavljen, posao koji jos nije krenuo vraca None i ne dira nista. Posao
    koji vrati STOP (npr. REVERT-FAIL) postavlja GRANICU pre nego sto vrati
    svog radnika: nijedan kasniji posao ne krece -- ni nad tom, nevracenom
    kopijom, ni nad drugima -- a raniji se zavrsavaju, jer ih potrosac jos cita.
    Potrosac mora da zatvori generator (contextlib.closing): tek tada se ceka
    kraj poslova u toku i izvrsilac se gasi."""
    slobodni = queue.Queue()
    for c in bazen:
        slobodni.put(c)
    granica = [len(poslovi)]
    brava = threading.Lock()

    def omot(i, fn):
        if _STOP.is_set() or i > granica[0]:
            return None
        c = slobodni.get()
        rez = None
        try:
            if _STOP.is_set() or i > granica[0]:
                return None
            rez = fn(c)
            return rez
        finally:
            # PRE vracanja radnika: posao koji ga sledeci uzme vec vidi granicu.
            if isinstance(rez, dict) and rez.get("stop"):
                with brava:
                    granica[0] = min(granica[0], i)
            slobodni.put(c)

    ex = concurrent.futures.ThreadPoolExecutor(max_workers=len(bazen))
    _AKTIVNI.append(ex)
    buduci = [ex.submit(omot, i, fn) for i, fn in enumerate(poslovi)]
    try:
        for f in buduci:
            # Ceka se u kratkim koracima: na Windows-u Ctrl+C ne prekida
            # beskonacno cekanje na bravu, pa bi glavna nit ostala gluva dok
            # radnici pale nove suite.
            while True:
                try:
                    r = f.result(timeout=0.5)
                    break
                except concurrent.futures.TimeoutError:
                    continue
            yield r
    except GeneratorExit:
        # Zatvaranje generatora (zip iscrpi poslove, ili potrosac stane) NIJE
        # prekid -- prekid je samo ako je neki posao jos u toku (finally).
        raise
    except BaseException:
        _STOP.set()
        raise
    finally:
        # Prekinut potrosac ne sme da ostavi poslove koji jos krecu.
        if any(not f.done() for f in buduci):
            _STOP.set()
        ex.shutdown(wait=True)
        # Uklanja se tek posle shutdown-a: prekinut join (drugi Ctrl+C) ostavlja
        # izvrsilac prijavljenim, pa ga rukovalac prekida i dalje vidi.
        _AKTIVNI.remove(ex)


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


def _baza_zelena(suite: str, testovi=None, src: str = None,
                 izlaz: str = None, prekini=None) -> tuple:
    """(je_zelena, opis). Tvrdi se, ne gleda se.

    Uz `testovi` je baza samo taj izbor -- tacno ono sto ce prolaz sa mutacijom
    pustiti. Izbor sa nula provera nije zelen nego prazan: nista nije mereno.
    `src` je kopija nad kojom se meri (D-6); izmedju prolaza je jednaka snimku.
    """
    try:
        r = _pusti_suite(*_run_suite(suite, testovi, src=src, izlaz=izlaz),
                         izlaz_dir=izlaz, prekini=prekini)
    except subprocess.TimeoutExpired:
        return False, f"{suite}: timeout"
    izlaz = r.stdout + r.stderr
    fatal = _fatal(izlaz)
    if fatal:
        return False, f"{suite}: EXCEL UBIJEN -- {fatal[:160]}"

    # Red je oznacen imenom suite-a, pa se dve suite u istom run-u ne mesaju.
    m = re.search(r"TESTS\s+" + re.escape(suite) + r": (\d+) ukupno, (\d+) palo",
                  izlaz)
    if not m:
        return False, f"{suite}: nema oznacenog reda TESTS u izlazu"
    if int(m.group(2)) != 0:
        return False, f"{suite}: {m.group(0)}"
    if testovi and int(m.group(1)) == 0:
        return False, f"{suite}: izbor bez ijedne provere ({m.group(0)})"
    return True, m.group(0)


# --- gde sidro zivi ---------------------------------------------------------
#
# Procedura u kojoj je sidro je jedino sto se o "kom kodu sabotaza pripada" moze
# utvrditi BEZ Excela, a dovoljno je grubo da ne laze: dve mutacije u istoj
# proceduri su po pravilu dva susedna uslova iste odluke, pa jedna lako obori
# tvrdnju druge. Fajl je za to previse krupan (svih 20 amb-pisac sabotaza zivi u
# dva fajla), a red previse sitan.
_PROC_KES = {}


def _procedura(fajl: str, sidro: str, sab, src: str = None) -> str:
    """'fajl::ImeProcedure' u kojoj sidro zivi, ili None ako se ne nalazi.

    None je ovde ZNACAJAN ishod, ne greska: sidro koje se ne nalazi je nalaz za
    `--proveri-sidra`, a ovde znaci "ne znam gde je" -- pa se takva sabotaza
    nikad ne grupise, nego meri sama.
    """
    kes = (src or SRC_VBA, fajl)
    if kes not in _PROC_KES:
        try:
            tekst = sab._procitaj(os.path.join(src or SRC_VBA, fajl))[0]
        except (OSError, UnicodeDecodeError, ValueError):
            _PROC_KES[kes] = None
        else:
            _PROC_KES[kes] = (tekst, sab._procedure_po_redu(tekst))
    if _PROC_KES[kes] is None:
        return None

    tekst, procedure = _PROC_KES[kes]
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


def _prolaz(clanovi: list, pre: str, grupno: bool, izbor=frozenset(),
            baze=None, ctx: dict = None) -> dict:
    """Jedan prolaz suite-a nad mutacijama svih clanova.

    Grupa od jednog clana JE pojedinacno merenje, pa je funkcija jedna. Razlika
    je samo u tome sta znaci neuspela primena: kod jednog clana je to nalaz
    (istrunulo sidro), a u grupi moze biti i SUDAR SIDARA -- partner je pregazio
    tekst na koji je ovo sidro zakaceno. Takav clan se zato ne prijavljuje kao
    APPLY-FAIL nego se vraca na pojedinacno merenje, gde pitanja o partnerima
    nema. Isto za clana koji u grupi nije oborio svoju tvrdnju.

    IZBOR (`izbor` = suite koje primaju `--test`): pusta se samo spisak testova
    clanova, a PRE prve mutacije isti spisak mora da bude zelen nad cistim
    izvorom (tacka 5 u docstring-u modula; `baze` pamti ishod po izboru). Ako
    nije, nista se ne primenjuje: grupa ide na pojedinacna merenja, a pojedinacan
    clan u "cela_suite" -- meri se celom suite-om, kao bez izbora.

    Clan koji u izboru NE obori svoju tvrdnju takodje ide u "cela_suite": test
    pusten sam moze da prodje mimo ciljane tvrdnje, pa presudu daje cela suite.

    `ctx` (D-6): mutacija ide u ctx["kopija"] (`--src`), run uvozi tu kopiju, a
    posle prolaza se kopija vraca iz ctx["snimak"]. src-vba se ne pise.

    Vraca {"stop": razlog ili "", "ishodi": [(stavka, priznato, stanje, crveno)],
           "ostatak": [(stavka, razlog)], "cela_suite": [(stavka, razlog)],
           "izbor": da li je prolaz isao izborom,
           "koren": potpis src-vba posle primene mutacija (ili "")}.
    """
    rez = {"stop": "", "ishodi": [], "ostatak": [], "cela_suite": [],
           "koren": ""}
    if not ctx or not _van_radnog_stabla(ctx["kopija"], ctx):
        rez["stop"] = ("STRAZA: nema kopije za mutacije, ili je kopija u "
                       "radnom stablu / snimku -- dokaz ne pise src-vba")
        return rez

    suite = clanovi[0]["suite"]
    testovi = None
    if suite in izbor:
        testovi = sorted({s["test"] for s in clanovi})
    rez["izbor"] = testovi is not None
    if testovi is not None:
        kljuc = (suite, tuple(testovi))
        baze = {} if baze is None else baze
        ok, opis = _baza_izbora(baze, kljuc, suite, testovi, ctx)
        if not ok:
            for s in clanovi:
                if grupno and len(clanovi) > 1:
                    rez["ostatak"].append((s, "baza izbora nije zelena -- %s" % opis))
                else:
                    rez["cela_suite"].append((s, opis))
            return rez

    primenjeni = []
    for s in clanovi:
        p = _pusti(sys.executable, ctx["alat"], s["ime"], "--src", ctx["kopija"])
        if p.returncode == 0:
            primenjeni.append(s)
        elif grupno:
            rez["ostatak"].append((s, "sidro se ne primenjuje uz partnere"))
        else:
            rez["ishodi"].append(
                (s, False, "APPLY-FAIL -- v. sabotaza.py --proveri-sidra", False))

    if not primenjeni:
        # Nista se nije PRIJAVILO kao primenjeno -- ali rc != 0 ne dokazuje da
        # fajl nije pisan (npr. ispis posle upisa koji pukne). Kopija se zato
        # vraca i ovde: mutacija koja ostane u njoj bila bi temelj svih sledecih
        # prolaza.
        sada = _vrati_kopiju(ctx, {s["fajl"] for s in clanovi}, pre)
        if sada != pre:
            rez["stop"] = ("REVERT-FAIL: kopija se ne vraca iz snimka (potpis %s)"
                           % sada)
        return rez

    # Radno stablo POSLE primene: dokaz tvrdi da ga ne pise, pa to i meri.
    rez["koren"] = _otisak()

    # Jedan primenjen clan je -- bez obzira na to kako je grupa planirana --
    # tacno pojedinacno merenje: u izvoru nema partnera, pa nema ni tvrdnje
    # koju bi partner mogao da obori. Zato se tako i ocenjuje (i ne vraca se
    # u red za ponovno merenje, sto bi bio isti prolaz drugi put).
    sam = len(primenjeni) == 1
    ocekivani = {s["kljuc"] for s in primenjeni}

    pali, greska = [], ""
    try:
        # Izbor je isti kao u bazi (svi clanovi, i oni ciji se sidro nije
        # primenilo): baza je merila TAJ spisak, pa se meri on.
        run = _pusti_suite(*_run_suite(suite, testovi, bez_markera=True,
                                       src=ctx["kopija"], izlaz=ctx.get("izlaz")),
                           izlaz_dir=ctx.get("izlaz"))
        izlaz = run.stdout + run.stderr
        fatal = _fatal(izlaz)
        if _STOP.is_set():
            # STOP ili Ctrl+C dok je run trajao: dete je verovatno prekinuto, pa
            # njegov izlaz (ili izostanak FAIL redova) ne kaze nista o tvrdnji.
            greska = "PREKID"
        elif fatal:
            # Prolaz nije stigao do kraja, pa njegovi FAIL redovi (ili njihov
            # izostanak) ne kazu nista o tvrdnji -- isto kao timeout.
            greska = "EXCEL UBIJEN -- %s" % fatal[:160]
        else:
            pali = _pali(izlaz, suite)
    except subprocess.TimeoutExpired:
        greska = "TIMEOUT suite"
    except KeyboardInterrupt:
        greska = "PREKID"
    finally:
        # Kopija se vraca iz SNIMKA i kad je run pukao: fajlovi svih clanova
        # (i onih ciji se sidro nije primenilo -- obnova je tada bez stete).
        # Snimak je istina, ne obrnuta zamena.
        sada = _vrati_kopiju(ctx, {s["fajl"] for s in clanovi}, pre)

    if sada != pre:
        rez["stop"] = "REVERT-FAIL: kopija se ne vraca iz snimka (potpis %s)" % sada
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
        if priznato:
            rez["ishodi"].append((s, True, stanje, bool(pali)))
        elif grupno and not sam:
            rez["ostatak"].append((s, stanje))
        elif testovi is not None:
            # IZBOR KOJI NE OBORI TVRDNJU NIJE PRESUDA. Test pusten sam moze da
            # prodje mimo ciljane tvrdnje (stanje koje mu inace ostavlja raniji
            # test ga vodi u drugu granu), pa bi "NIJE DOKAZANO" ovde bilo lazno
            # crveno o katalogu. Presudu daje cela suite, kao bez --izbor.
            rez["cela_suite"].append((s, "u izboru: %s" % stanje))
        else:
            rez["ishodi"].append((s, False, stanje, bool(pali)))
    return rez


def _verdikt(ok: bool, grupno_priznatih: int, izborom_priznatih: int,
             snimak: str = "") -> str:
    """Zavrsni red dokaza. Svaki rezim koji meri manje nosi svoje ime.

    `grupno`: mutacije grupe su pustene ZAJEDNO (tacka 4). `izbor`: ciljani test
    je pusten SAM, ne u punom redosledu suite-a (tacka 5). `snimak`: radno
    stablo se menjalo tokom dokaza, pa dokaz vazi za SNIMAK sa tim potpisom, ne
    za ono sto sada stoji u src-vba (tacka 2). Golo `DOKAZANO` znaci samo jedno:
    svaka sabotaza pojedinacno, u celoj suiti, nad izvorom koji i dalje stoji.
    """
    if not ok:
        return "NIJE DOKAZANO"
    oznake = [ime for ime, n in (("grupno", grupno_priznatih),
                                 ("izbor", izborom_priznatih)) if n]
    if snimak:
        oznake.append("snimak %s" % snimak)
    return "DOKAZANO (%s)" % ", ".join(oznake) if oznake else "DOKAZANO"


def _izaberi(sab, filter_: list, src: str = None) -> list:
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
            "proc": _procedura(fajl, sidro, sab, src),
        })
    return stavke


_BAZEN = []        # radnici tekuceg dokaza -- za ciscenje pri prekidu
_AKTIVNI = []      # izvrsioci u toku -- prekid ceka njih PRE ciscenja
_UNAPRED = []      # baze unapred u toku -- prekid ih OTKAZUJE pre cekanja


class _BazeUnapred:
    """Pune baze cele suite UNAPRED, na rezervisanom radniku (R-4).

    Uz --izbor se puna baza trazi tek za sabotaze koje izbor ne obori (faza
    3), pa je bila na kraju dokaza: puna suite za bazu, pa puna suite za
    mutaciju -- jedna za drugom, dok ostali radnici stoje. Ovde krece na
    pocetku, na radniku koga faze za to vreme ne koriste (ukupno i dalje N
    Excel-a). Baze idu JEDNA PO JEDNA: pisu marker zelenog, a drugih pisaca
    markera tada nema (izbor i mutacije ga ne pisu).

    Ne zatreba li faza 3, baza se OTKAZUJE: njen Excel se ubija po excel.pid
    (samo uz isto vreme nastanka), pa mali dokaz ne ceka punu suite koju
    nista ne koristi."""

    def __init__(self, suite: list, ctx: dict, ubij=None):
        self.suite = list(suite)
        self.ctx = ctx
        self.rez = {}
        self._brava = threading.Lock()
        self._otkazi = threading.Event()
        self._ubij = ubij or (lambda: _ubij_zaostali(ctx.get("izlaz")))
        self._ex = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        _AKTIVNI.append(self._ex)
        _UNAPRED.append(self)
        self._f = self._ex.submit(self._radi)

    def _radi(self) -> None:
        for suite in self.suite:
            if self._otkazi.is_set() or _STOP.is_set():
                return
            rez = _baza_zelena(suite, src=self.ctx["kopija"],
                               izlaz=self.ctx.get("izlaz"), prekini=self._otkazi)
            with self._brava:
                self.rez[suite] = rez

    def gotovo(self) -> bool:
        return self._f.done()

    def gotove(self) -> set:
        with self._brava:
            return set(self.rez)

    def sacekaj(self, potrebne=None) -> dict:
        """Sacekaj baze koje faza 3 TREBA (u koracima -- Ctrl+C mora da stigne)
        i vrati gotove. Cim su potrebne gotove, ostatak se otkazuje: faza 3 ne
        ceka punu suite koja joj ne treba."""
        potrebne = set(self.suite if potrebne is None else potrebne) & set(self.suite)
        while not self._f.done():
            if potrebne <= self.gotove():
                self.otkazi()
                break
            try:
                self._f.result(timeout=0.5)
            except concurrent.futures.TimeoutError:
                continue
        if self._f.done() and not self._otkazi.is_set():
            self._f.result()            # izuzetak iz baze ide gore, kao do sada
        self._zatvori()
        with self._brava:
            return dict(self.rez)

    def otkazi(self) -> None:
        """Baze nisu potrebne: ne pokreci sledecu, a tekucu prekini."""
        self._otkazi.set()
        while not self._f.done():
            self._ubij()
            try:
                self._f.result(timeout=0.5)
            except concurrent.futures.TimeoutError:
                continue
            except Exception:               # noqa: BLE001 -- otkazana baza nije nalaz
                break
        self._zatvori()

    def _zatvori(self) -> None:
        self._ex.shutdown(wait=True)
        if self._ex in _AKTIVNI:
            _AKTIVNI.remove(self._ex)
        if self in _UNAPRED:
            _UNAPRED.remove(self)


def _serijska_traka(stavke: list, radnika: int) -> str:
    """Razlog odbijanja ili "": uz radnike smeju samo PARALELNE_SUITE."""
    if radnika <= 1:
        return ""
    van = sorted({s["suite"] for s in stavke} - PARALELNE_SUITE)
    if van:
        return ("--radnika %d: suite %s nisu provereno nezavisne (mreza, deljeno "
                "stanje) -- pusti ih sa --radnika 1" % (radnika, ", ".join(van)))
    return ""


def _spoji_istoriju(bazen: list, radnika: int, cilj: str = None) -> str:
    """Redovi istorije iz izlaza radnika -> tests/run_vba_istorija.jsonl.

    Svaki run_vba pise SVOJ izlaz (bez deljenog upisa); ovde ih spaja jedan
    pisac, na kraju, pod bravom tog fajla. Vraca "" ili razlog."""
    cilj = cilj or os.path.join(ROOT, "tests", "run_vba_istorija.jsonl")
    redovi = []
    for c in bazen:
        put = os.path.join(c.get("izlaz") or "", "run_vba_istorija.jsonl")
        if not os.path.isfile(put):
            continue
        with open(put, encoding="utf-8") as fh:
            for red in fh:
                red = red.strip()
                if not red:
                    continue
                try:
                    zapis = json.loads(red)
                except ValueError:
                    continue
                zapis["dokaz_radnika"] = radnika
                redovi.append(json.dumps(zapis, ensure_ascii=True))
    if not redovi:
        return ""
    try:
        brava = _run_vba_modul()._zakljucaj(
            os.path.join(os.path.dirname(cilj), ".run_vba.lock"), cekaj=5.0)
    except Exception:                       # noqa: BLE001
        brava = None
    if brava is None:
        # tests/ drzi rucni run_vba: redovi idu u RESERVNI fajl pored, ne nestaju.
        rezerva = os.path.join(os.path.dirname(cilj),
                               "run_vba_istorija.dokaz_%d.jsonl" % os.getpid())
        try:
            with open(rezerva, "a", encoding="utf-8") as fh:
                fh.write("\n".join(redovi) + "\n")
        except OSError as exc:
            return "ISTORIJA NIJE SACUVANA: %s" % exc
        _obrisi_istoriju_radnika(bazen)
        return "istorija je u %s (tests/ je bio zauzet -- rucni run_vba?)" % rezerva
    try:
        with open(cilj, "a", encoding="utf-8") as fh:
            fh.write("\n".join(redovi) + "\n")
    except OSError as exc:
        return "ISTORIJA NIJE SACUVANA: %s" % exc
    finally:
        _run_vba_modul()._pusti_bravu(brava)
    _obrisi_istoriju_radnika(bazen)
    return ""


def _obrisi_istoriju_radnika(bazen: list) -> None:
    """Spojeni redovi se brisu iz izlaza radnika, pa drugi poziv ne udvaja."""
    for c in bazen:
        try:
            os.remove(os.path.join(c.get("izlaz") or "", "run_vba_istorija.jsonl"))
        except OSError:
            pass


def _zavrsi_istoriju(bazen: list, radnika: int) -> bool:
    """Spoji istoriju i javi ishod. False = redovi NISU sacuvani (radni ostaje)."""
    razlog = _spoji_istoriju(bazen, radnika)
    if razlog:
        print(razlog, file=sys.stderr)
    return not razlog.startswith("ISTORIJA NIJE SACUVANA")


def _izolacija(stavke: list, izbor: set, src: str = None, bazen=None) -> list:
    """[(suite, test, opis)] za ciljane testove koji SAMI ne prolaze.

    Bez mutacija: svaki razlicit ciljani test se pusti sam nad cistim izvorom.
    Crven test nije kvar izbora nego nalaz o testu -- zavisi od stanja koje mu
    ostavlja raniji test u suite-u. `--izbor` ga ionako meri celom suite-om, a
    ovde se nabraja da se zna koliko testova tako zavisi od redosleda.
    """
    crveni = []
    vidjeni, jedinstveni = set(), []
    for s in stavke:
        kljuc = (s["suite"], s["test"])
        if s["suite"] not in izbor or kljuc in vidjeni:
            continue
        vidjeni.add(kljuc)
        jedinstveni.append(s)
    bazen = bazen or [{"kopija": src, "izlaz": None}]
    with contextlib.closing(_izvrsi([(lambda c, s=s: _baza_zelena(s["suite"], [s["test"]],
                                                    src=c["kopija"], izlaz=c["izlaz"]))
                       for s in jedinstveni], bazen)) as tok:
        for s, rez in zip(jedinstveni, tok):
            if rez is None:
                break
            ok, opis = rez
            print("%-6s %-28s %-50s %s" % ("OK" if ok else "CRVEN", s["suite"],
                                           s["test"], "" if ok else opis), flush=True)
            if not ok:
                crveni.append((s["suite"], s["test"], opis))
    print("\nizolacija: %d testova, %d ne prolazi samo" % (len(vidjeni), len(crveni)))
    return crveni


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
    ap.add_argument("--izbor", action="store_true",
                    help="pusti samo ciljane testove (run_vba --test) umesto cele "
                         "suite; baza se tvrdi nad istim izborom")
    ap.add_argument("--izolacija", action="store_true",
                    help="bez mutacija: pusti svaki ciljani test SAM i nabroji "
                         "one koji ne prolaze")
    ap.add_argument("--radnika", type=int, default=1, metavar="N",
                    help="N paralelnih Excel radnika (svaki sa svojom kopijom izvora "
                         "i svojim izlazom); 1 = serijski, kao do sada")
    ap.add_argument("--razmak", type=float, default=1.0, metavar="S",
                    help="najmanje S sekundi izmedju dva run_vba poziva (uz "
                         "--radnika > 1); podrazumevano 1.0")
    ap.add_argument("--self-test", action="store_true",
                    help="dokazi da pravila grupisanja, ocena i orkestracija prolaza grizu")
    args = ap.parse_args(argv)

    if args.self_test:
        return _self_test()
    if args.grupe < 1:
        print("--grupe mora biti >= 1", file=sys.stderr)
        return 2
    if args.radnika < 1:
        print("--radnika mora biti >= 1", file=sys.stderr)
        return 2
    _STOP.clear()
    _START.update(poslednji=0.0, ponovljeno=0,
                  razmak=args.razmak if args.radnika > 1 else 0.0)

    sab = _modul_sabotaza()
    poznati_spisak = getattr(sab, "POZNATI_NALAZI_DOKAZ", {})
    stavke = _izaberi(sab, args.filter)

    if not stavke:
        print("filter ne pogadja nijednu sabotazu", file=sys.stderr)
        return 2

    if args.plan:
        return _plan(stavke, args.grupe)

    problem = _serijska_traka(stavke, args.radnika)
    if problem:
        print(problem, file=sys.stderr)
        return 2

    # --- D-6: snimak i kopija PRE prve baze --------------------------------
    #
    # Sve od ovog trenutka meri SNIMAK: baze i mutacije idu nad kopijom, a
    # src-vba se samo cita (potpis pre/tokom/posle). Procedure sidara se citaju
    # iz snimka, da plan grupa i mutacija gledaju isti tekst.
    global _RADNI
    ctx, problem = _pripremi_radni(sab, radnika=args.radnika)
    _RADNI = ctx["radni"]
    bazen = ctx.get("bazen") or [ctx]
    _BAZEN[:] = bazen
    if problem:
        print("STOP: " + problem, file=sys.stderr)
        shutil.rmtree(ctx["radni"], ignore_errors=True)
        return 2
    stavke = _izaberi(sab, args.filter, ctx["snimak"])

    izbor = _suite_izbor() if (args.izbor or args.izolacija) else set()
    if args.izolacija:
        try:
            return 1 if _izolacija(stavke, izbor, ctx["kopija"], bazen) else 0
        finally:
            if _zavrsi_istoriju(bazen, args.radnika):
                shutil.rmtree(ctx["radni"], ignore_errors=True)

    print("sabotaza: %d%s" % (len(stavke), " (izbor: samo ciljani testovi)"
                              if izbor else ""), flush=True)

    # --- KAPIJA: baza mora biti zelena pre prve mutacije --------------------
    #
    # Uz izbor se baza suite-a koja prima `--test` tvrdi po IZBORU, u svakom
    # prolazu (v. `_prolaz`), pa se ovde tvrdi samo za suite bez izbora.
    potrebne = sorted({s["suite"] for s in stavke} | {SUITE_ALL})
    if izbor:
        potrebne = sorted({s["suite"] for s in stavke} - izbor)
    cele_baze = {}
    for suite in potrebne:
        ok, opis = _baza_zelena(suite, src=ctx["kopija"], izlaz=ctx["izlaz"])
        cele_baze[suite] = (ok, opis)
        print("BAZNO: %s" % opis, flush=True)
        if not ok:
            print("STOP: baza nije zelena. Dokaz bi merio crveno koje sabotaza "
                  "nije izazvala. Kopija ostaje u %s" % ctx["radni"],
                  file=sys.stderr)
            _zavrsi_istoriju(bazen, args.radnika)
            return 2

    pre = _otisak(ctx["snimak"])
    # src-vba je bio jednak snimku u trenutku kopiranja (tvrdi _pripremi_radni);
    # izmena tokom baza zato mora da se vidi -- ne uzima se nov uzorak.
    koren_pre = pre
    koren_tokom = set()
    print("potpis izvora: %s (snimak; src-vba se ne pise)" % pre, flush=True)

    # Baze cele suite UNAPRED (v. _BazeUnapred): samo uz radnike i izbor, na
    # poslednjem radniku; faze ga ne koriste dok baze traju.
    unapred = None
    if args.radnika > 1 and izbor:
        moguce = sorted(({s["suite"] for s in stavke} & izbor) - set(cele_baze))
        if moguce:
            unapred = _BazeUnapred(moguce, bazen[-1])
            print("BAZE UNAPRED: %s (radnik r%d) -- za sabotaze koje izbor ne "
                  "obori; otkazuju se ako ne zatrebaju"
                  % (", ".join(moguce), bazen[-1]["id"]), flush=True)

    def bazen_faze():
        """Radnici za fazu: bez rezervisanog dok baze unapred traju."""
        if unapred is not None and not unapred.gotovo():
            return bazen[:-1]
        return bazen

    crvenih, lose, poznati = 0, [], []
    grupno_priznatih, izborom_priznatih, stop = 0, 0, ""

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
    baze_izbora = {}
    cela_suite = []
    if args.grupe > 1:
        print("grupisanje: %d prolaza nad %d sabotaza (najvise %d po grupi); "
              "clan koji u grupi ne obori SVOJU tvrdnju meri se posle sam"
              % (len(grupe), len(stavke), args.grupe), flush=True)

    if args.radnika > 1:
        print("radnika: %d (svaki u svom Excel-u, nad svojom kopijom izvora); "
              "ishodi se ispisuju u redosledu kataloga" % args.radnika, flush=True)

    def posao(clanovi, grupno, izb=frozenset(), baz=None):
        return lambda c: _prolaz(clanovi, pre, grupno, izb, baz, c)

    with contextlib.closing(_izvrsi([posao(g, len(g) > 1, izbor, baze_izbora)
                                     for g in grupe], bazen_faze())) as tok:
        for i, (g, ishod) in enumerate(zip(grupe, tok), 1):
            if ishod is None:
                stop = stop or "PREKID"
                break
            if len(g) > 1:
                print("GRUPA %d/%d  %s  clanova %d"
                      % (i, len(grupe), g[0]["suite"], len(g)), flush=True)
            koren_tokom.add(ishod.get("koren") or koren_pre)
            cela_suite.extend(ishod["cela_suite"])
            if len(g) > 1:
                grupno_priznatih += sum(1 for _s, p, _st, _c in ishod["ishodi"] if p)
            if ishod.get("izbor"):
                izborom_priznatih += sum(1 for _s, p, _st, _c in ishod["ishodi"] if p)
            upisi(ishod["ishodi"], "  " if len(g) > 1 else "")
            for s, razlog in ishod["ostatak"]:
                print("  %-46s -> ponavlja se sam (%s)" % (s["ime"], razlog),
                      flush=True)
                ponavljaju.append(s)
            if ishod["stop"]:
                stop = ishod["stop"]
                _STOP.set()
                break

    if ponavljaju and not stop:
        print("PONAVLJA SE SAM: %d" % len(ponavljaju), flush=True)
        with contextlib.closing(_izvrsi([posao([s], False, izbor, baze_izbora)
                                         for s in ponavljaju], bazen_faze())) as tok:
            for s, ishod in zip(ponavljaju, tok):
                if ishod is None:
                    stop = stop or "PREKID"
                    break
                koren_tokom.add(ishod.get("koren") or koren_pre)
                cela_suite.extend(ishod["cela_suite"])
                if ishod.get("izbor"):
                    izborom_priznatih += sum(1 for _s, p, _st, _c in ishod["ishodi"] if p)
                upisi(ishod["ishodi"])
                for s2, razlog in ishod["ostatak"]:
                    # Pojedinacno merenje nema partnere, pa nema ni sta da vrati u
                    # red -- ali ako ga ikad vrati, sabotaza bi tiho ispala iz
                    # imenioca. Zato je to nalaz, a ne tiho preskakanje.
                    lose.append((s2["ime"], "NIJE IZMERENA: %s" % razlog))
                    print("%-46s NIJE IZMERENA: %s" % (s2["ime"], razlog), flush=True)
                if ishod["stop"]:
                    stop = ishod["stop"]
                    _STOP.set()
                    break

    # Test koji SAM ne prolazi meri se celom suite-om -- uz bazu CELE suite,
    # koja se tvrdi jednom po suite-u. Baze idu SERIJSKI, pre radnika: pisu
    # marker zelenog (vba_gate), a on nije za istovremene upise.
    if cela_suite and not stop:
        print("MERI SE CELOM SUITE-OM: %d (ciljani test sam ne prolazi, ili izbor "
              "nije oborio tvrdnju -- v. --izolacija)" % len(cela_suite), flush=True)
        for s, razlog in cela_suite:
            print("  %-46s <- %s" % (s["ime"], razlog), flush=True)
        if unapred is None:
            # Serijski put (i radnici bez baza unapred): baza, pa mutacije.
            for suite in sorted({s["suite"] for s, _r in cela_suite} - set(cele_baze)):
                ok_opis = _baza_zelena(suite, src=ctx["kopija"], izlaz=ctx["izlaz"])
                cele_baze[suite] = ok_opis
                print("BAZNO: %s" % cele_baze[suite][1], flush=True)
            merljive = []
            for s, _razlog in cela_suite:
                ok, opis = cele_baze[s["suite"]]
                if not ok:
                    lose.append((s["ime"], "NIJE IZMERENA: ni izbor ni cela suite "
                                           "nisu zeleni -- %s" % opis))
                    print("%-46s NIJE IZMERENA: baza cele suite nije zelena"
                          % s["ime"], flush=True)
                    continue
                merljive.append(s)
            with contextlib.closing(_izvrsi([posao([s], False) for s in merljive],
                                            bazen)) as tok:
                for s, ishod in zip(merljive, tok):
                    if ishod is None:
                        stop = stop or "PREKID"
                        break
                    koren_tokom.add(ishod.get("koren") or koren_pre)
                    upisi(ishod["ishodi"])
                    if ishod["stop"]:
                        stop = ishod["stop"]
                        _STOP.set()
                        break
        else:
            # RANA MUTACIJA (R-4): mutacije cele suite krecu ODMAH, uz bazu
            # unapred koja jos traje, a baza se ceka tek pri oceni. Mutacija se
            # priznaje SAMO uz zelenu bazu -- crvena baza i dalje daje NIJE
            # IZMERENA -- pa je tvrdnja ista kao serijski; menja se raspored, ne
            # tumacenje. Kraj dokaza je tako max(baza, mutacija), ne zbir.
            potrebne_3 = {s["suite"] for s, _r in cela_suite}
            for suite in sorted(potrebne_3 - set(unapred.suite) - set(cele_baze)):
                ok_opis = _baza_zelena(suite, src=ctx["kopija"], izlaz=ctx["izlaz"])
                cele_baze[suite] = ok_opis
                print("BAZNO: %s" % cele_baze[suite][1], flush=True)
            sve = [s for s, _r in cela_suite]
            with contextlib.closing(_izvrsi([posao([s], False) for s in sve], bazen_faze())) as tok:
                for s, ishod in zip(sve, tok):
                    if ishod is None:
                        stop = stop or "PREKID"
                        break
                    koren_tokom.add(ishod.get("koren") or koren_pre)
                    # STOP ne ceka punu suite baze: dokaz je vec nevazeci, pa se
                    # uzima samo baza koja je VEC gotova.
                    if unapred is not None and (not ishod["stop"] or unapred.gotovo()):
                        for suite, ok_opis in sorted(unapred.sacekaj(potrebne_3).items()):
                            if suite not in potrebne_3:
                                continue
                            cele_baze[suite] = ok_opis
                            print("BAZNO (unapred): %s" % ok_opis[1], flush=True)
                        unapred = None
                    ok, opis = cele_baze.get(s["suite"],
                                             (False, "baza cele suite nije izmerena "
                                                     "(dokaz je stao pre nje)"))
                    if ok:
                        upisi(ishod["ishodi"])
                    else:
                        # Mutacija je izmerena, ali bez zelene baze ne kaze nista.
                        lose.append((s["ime"], "NIJE IZMERENA: ni izbor ni cela suite "
                                               "nisu zeleni -- %s" % opis))
                        print("%-46s NIJE IZMERENA: %s"
                              % (s["ime"], "baza cele suite nije zelena"
                                 if s["suite"] in cele_baze
                                 else "dokaz je stao pre baze cele suite"), flush=True)
                    if ishod["stop"]:
                        stop = ishod["stop"]
                        _STOP.set()
                        break

    if stop:
        if stop.startswith("REVERT-FAIL"):
            print("STOP: kopija izvora nije vracena u stanje snimka. Sve mereno "
                  "posle ovoga islo bi nad pokvarenim kodom. src-vba nije diran; "
                  "kopija ostaje u %s" % ctx["radni"], file=sys.stderr)
        elif stop.startswith("STRAZA"):
            print("STOP: " + stop, file=sys.stderr)

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

    if unapred is not None:
        # Faza 3 nije trebala (ili je dokaz stao): baza unapred se otkazuje.
        vec_gotove = unapred.gotovo()
        unapred.otkazi()
        print("BAZE UNAPRED %s" % (
            "otkazane -- dokaz je stao" if stop and not vec_gotove
            else "nisu zatrebale -- %s" % ("bile su vec gotove" if vec_gotove
                                           else "otkazane")), flush=True)
        unapred = None

    posle = _otisak(ctx["snimak"])
    koren_menjan = _koren_menjan(pre, koren_tokom, _otisak())
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
    print("izvor (snimak) pre/posle: %s / %s -> %s"
          % (pre, posle, "IDENTICAN" if pre == posle else "RAZLIKA!"))
    print("radno stablo src-vba pre/tokom/posle: %s"
          % ("PROMENJENO -- dokaz je merio snimak %s, ne sadasnji src-vba" % pre
             if koren_menjan else "IDENTICNO (dokaz ga nije pisao)"))
    if args.radnika > 1 or _START["ponovljeno"]:
        print("radnika: %d; ponovljenih poziva: %d (FATAL pre ijedne suite -- "
              "Excel ili uvoz nisu prosli; meri se drugi pokusaj)"
              % (args.radnika, _START["ponovljeno"]))
    if izbor:
        print("izbor: baza i mutacija nad istim spiskom testova; celom suite-om "
              "mereno %d (test sam ne prolazi, ili izbor nije oborio tvrdnju)"
              % len(cela_suite))
    if izborom_priznatih:
        print("izborom izmereno: %d / %d -- ciljani test je pusten SAM, ne u "
              "punom redosledu suite-a; dokaz u redosledu je isti poziv bez "
              "--izbor" % (izborom_priznatih, len(stavke)))
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
    print("=== %s ===" % _verdikt(ok, grupno_priznatih, izborom_priznatih,
                                 pre if koren_menjan else ""))
    sacuvana = _zavrsi_istoriju(bazen, args.radnika)
    if stop or not sacuvana:
        print("radni folder ostaje: %s" % ctx["radni"], file=sys.stderr)
    else:
        shutil.rmtree(ctx["radni"], ignore_errors=True)
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


# Lazni radni folder: putanje koje ne postoje, van src-vba i van snimka.
_LAZNI_CTX = {"radni": os.path.join(os.sep, "nema", "dokaz_x"),
              "snimak": os.path.join(os.sep, "nema", "dokaz_x", "snimak", "src-vba"),
              "kopija": os.path.join(os.sep, "nema", "dokaz_x", "r1", "src-vba"),
              "alat": os.path.join(os.sep, "nema", "dokaz_x", "alat", "sabotaza.py")}


def _lazni_prolaz(clanovi, izlaz_suite, grupno=True, pukli=(), potpis="P",
                  timeout=False, izbor=frozenset(), baze=None, izlaz_baze="",
                  ctx=None, prekid=False):
    """Pokreni `_prolaz` bez Excela. Vraca (rez, pozivi).

    Pravila grupisanja i ocena su ciste funkcije i mere se direktno, ali
    ORKESTRACIJA nije: primeni sve (u kopiju) -> pusti suite TACNO JEDNOM ->
    obnovi kopiju iz snimka -> oceni, pa clan koji nije priznat vrati na
    pojedinacno merenje. Bez ovoga bi
    taj redosled prvi put bio izvrsen nad pravim Excelom, gde jedan prolaz traje
    minutima i gde greska u njemu izgleda kao nalaz nad VBA kodom.
    """
    pozivi = []
    stanje = {"mutirano": False}

    def lazni_pusti(*a, **kw):
        pozivi.append(tuple(a[1:]))
        if a[1].endswith("run_vba.py"):
            # Run pre ijedne primenjene mutacije je BAZA, i dobija svoj izlaz --
            # tako se vidi da se baza meri nad cistim izvorom, a ne posle.
            if not stanje["mutirano"]:
                return _LazniOdgovor(0, izlaz_baze)
            if timeout:
                raise subprocess.TimeoutExpired(a, 1)
            if prekid:
                raise KeyboardInterrupt
            return _LazniOdgovor(0, izlaz_suite)
        # Primena sabotaze. Vracanja vise nema (D-6): kopija se obnavlja iz
        # snimka, pa se obnova belezi u isti niz poziva, da se vidi redosled.
        rc = 2 if a[2] in pukli else 0
        stanje["mutirano"] = stanje["mutirano"] or rc == 0
        return _LazniOdgovor(rc)

    def lazna_obnova(_ctx, fajlovi=None):
        pozivi.append(("<obnova>", "<cela>" if fajlovi is None
                       else tuple(sorted(fajlovi))))
        stanje["mutirano"] = False

    # Potpis ZAVISI OD PUTANJE: `potpis` vazi samo za kopiju, sve ostalo
    # (src-vba, snimak) je "P". Provera koja bi merila pogresno stablo -- npr.
    # src-vba umesto kopije -- tako bi videla "P" i propustila nevracenu kopiju.
    kopija = (ctx or _LAZNI_CTX)["kopija"]
    pravi = _pusti, _otisak, _obnovi
    globals()["_pusti"] = lazni_pusti
    globals()["_otisak"] = lambda put=None: potpis if put == kopija else "P"
    globals()["_obnovi"] = lazna_obnova
    try:
        rez = _prolaz(clanovi, "P", grupno, izbor, baze, ctx or _LAZNI_CTX)
    finally:
        globals()["_pusti"], globals()["_otisak"], globals()["_obnovi"] = pravi
    return rez, pozivi


def _self_test_d6(prolaz, suita) -> list:
    """D-6 u oba smera: kopija, obnova, straza, priprema snimka, ozicenje main-a."""
    nalazi = []

    def tvrdi(uslov, opis):
        if not uslov:
            nalazi.append(opis)

    K = _LAZNI_CTX["kopija"]
    g2 = [_s("a", test="T_A", kljuc="T_A", tvrdnja="razlog a", proc="modA.bas::P"),
          _s("b", fajl="modB.bas", test="T_B", kljuc="T_B", tvrdnja="razlog b",
             proc="modB.bas::Q")]
    OBA = "FAIL T_A -- razlog a\nFAIL T_B -- razlog b\n"
    _rez, pozivi = prolaz(g2, OBA)
    primene = [p for p in pozivi if p and p[0].endswith("sabotaza.py")]
    tvrdi(len(primene) == 2 and all(p[0] == _LAZNI_CTX["alat"]
                                    and p[-2:] == ("--src", K) for p in primene),
          "D-6: sabotaza se ne primenjuje zamrznutim katalogom u kopiju (--src)")
    tvrdi(not any("--vrati" in p for p in pozivi),
          "D-6: dokaz zove --vrati (radno stablo) umesto obnove iz snimka")
    tvrdi(suita(pozivi) and all(r[-2:] == ("--src", K) for r in suita(pozivi)),
          "D-6: run_vba ne uvozi kopiju (--src), nego src-vba")
    tvrdi([p for p in pozivi if p[0] == "<obnova>"]
          == [("<obnova>", ("modA.bas", "modB.bas"))],
          "D-6: obnova ne vraca bas fajlove clanova prolaza")

    for opis, kw in (("timeout", {"timeout": True}), ("Ctrl+C", {"prekid": True})):
        _rez, pozivi = prolaz([g2[0]], OBA, grupno=False, **kw)
        tvrdi(pozivi and pozivi[-1] == ("<obnova>", ("modA.bas",)),
              "D-6: kopija se ne vraca iz snimka posle %s" % opis)

    # rc != 0 ne dokazuje da fajl nije pisan: kopija se vraca i kad se NIJEDNA
    # primena ne prijavi, a nevracena kopija i tada zaustavlja dokaz.
    rez, pozivi = prolaz(g2, OBA, pukli=("a", "b"))
    tvrdi(not suita(pozivi) and pozivi[-1] == ("<obnova>", ("modA.bas", "modB.bas"))
          and not rez["stop"],
          "D-6: kopija se ne vraca kad nijedna primena nije prijavljena")
    rez, _p = prolaz(g2, OBA, pukli=("a", "b"), potpis="DRUGI")
    tvrdi(rez["stop"].startswith("REVERT-FAIL"),
          "D-6: nevracena kopija bez prijavljene primene ne zaustavlja dokaz")

    # Baza izbora i izolacija takodje mere KOPIJU.
    IZB = frozenset({SUITE_ALL})
    _rez, pozivi = prolaz(g2, OBA, izbor=IZB, baze={},
                          izlaz_baze="TESTS   RunAllTests: 2 ukupno, 0 palo\n")
    runovi = suita(pozivi)
    tvrdi(len(runovi) == 2 and all(r[-2:] == ("--src", K) for r in runovi),
          "D-6: baza izbora ne uvozi kopiju (--src), nego src-vba")
    izol = []
    pravi_pusti = _pusti
    globals()["_pusti"] = lambda *a, **kw: (izol.append(a[1:]) or _LazniOdgovor(
        0, "TESTS   RunAllTests: 1 ukupno, 0 palo\n"))
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            _izolacija([_s("a")], {SUITE_ALL}, K)
    finally:
        globals()["_pusti"] = pravi_pusti
    tvrdi(izol and all(r[-2:] == ("--src", K) for r in izol),
          "D-6: izolacija ne uvozi kopiju (--src), nego src-vba")

    # Radno stablo se poredi sa SNIMKOM (izmena tokom baza mora da se vidi).
    for opis, tokom, posle, ocekivano in (
            ("nepromenjeno", {"S"}, "S", False),
            ("promena tokom prolaza", {"S", "X"}, "S", True),
            ("promena tokom baza (svaki uzorak je vec nov)", {"X"}, "X", True),
            ("promena posle poslednjeg prolaza", {"S"}, "X", True),
            ("prolaz bez primene (prazan uzorak)", {""}, "S", False)):
        tvrdi(_koren_menjan("S", tokom, posle) == ocekivano,
              "D-6: radno stablo -- %s -- daje %r" % (opis, not ocekivano))

    # Straza: cilj mutacije nikad nije src-vba ni snimak.
    tvrdi(not _van_radnog_stabla(SRC_VBA)
          and not _van_radnog_stabla(os.path.join(SRC_VBA, "pod")),
          "D-6: straza pusta src-vba (ili folder ispod njega) kao cilj mutacije")
    tvrdi(not _van_radnog_stabla(_LAZNI_CTX["snimak"], _LAZNI_CTX),
          "D-6: straza pusta snimak kao cilj mutacije")
    tvrdi(_van_radnog_stabla(K, _LAZNI_CTX),
          "D-6: straza odbija ispravnu kopiju")
    rez, pozivi = prolaz(g2, OBA, ctx=dict(_LAZNI_CTX,
                                           kopija=os.path.join(SRC_VBA, "pod")))
    tvrdi(rez["stop"].startswith("STRAZA") and not pozivi,
          "D-6: kopija u radnom stablu ipak dobija mutaciju ili run")
    rez, pozivi = prolaz(g2, OBA, ctx=dict(_LAZNI_CTX, kopija=_LAZNI_CTX["snimak"]))
    tvrdi(rez["stop"].startswith("STRAZA") and not pozivi,
          "D-6: snimak kao kopija ipak dobija mutaciju ili run")

    tvrdi(_run_suite(SUITE_ALL, ["T_A"], True, "/k")[-2:] == ("--src", "/k")
          and "--src" not in _run_suite(SUITE_ALL),
          "D-6: _run_suite ne predaje --src na kraju (ili ga predaje bez kopije)")

    # Priprema nad MALIM korenom, ne nad src-vba (self-test ide kroz hook).
    class _LazniSab:
        def __init__(self, zatecene):
            self._z = zatecene

        def zatecene(self, _src):
            return list(self._z)

    koren = tempfile.mkdtemp(prefix="dokaz_st_")
    try:
        for ime in ("modA.bas", "modB.bas"):
            with open(os.path.join(koren, ime), "w", newline="\r\n") as fh:
                fh.write('Attribute VB_Name = "%s"\n' % ime[:-4])
        pre_koren = _otisak(koren)
        ctx, problem = _pripremi_radni(_LazniSab([]), koren)
        try:
            tvrdi(not problem
                  and _otisak(ctx["snimak"]) == _otisak(ctx["kopija"]) == pre_koren
                  and os.path.isfile(ctx["alat"]),
                  "D-6: priprema ne daje snimak i kopiju jednake izvoru (%s)" % problem)
            # Zamrznut katalog bez --src gadja <koren alata>/src-vba (isto
            # pravilo kao ROOT/SRC_VBA u sabotaza.py) -- a taj folder NE SME da
            # postoji: inace bi zaboravljen --src pokvario snimak ili kopiju.
            podrazumevani = os.path.join(
                os.path.dirname(os.path.dirname(ctx["alat"])), "src-vba")
            tvrdi(not os.path.exists(podrazumevani),
                  "D-6: zamrznut katalog bez --src gadja postojeci folder (%s)"
                  % podrazumevani)
            tvrdi(_van_radnog_stabla(ctx["kopija"], ctx)
                  and not _u_stablu(ctx["kopija"], ctx["snimak"]),
                  "D-6: kopija je u snimku")
            # PRAVA obnova nad pravim folderom: po fajlu, pa cela posle
            # obrisanog fajla. Snimak se pri tome ne sme promeniti.
            snimak_pre = _otisak(ctx["snimak"])
            with open(os.path.join(ctx["kopija"], "modA.bas"), "a") as fh:
                fh.write("' mutacija\n")
            sada_fajl = _vrati_kopiju(ctx, {"modA.bas"}, snimak_pre)
            os.remove(os.path.join(ctx["kopija"], "modB.bas"))
            with open(os.path.join(ctx["kopija"], "modA.bas"), "a") as fh:
                fh.write("' mutacija\n")
            sada_cela = _vrati_kopiju(ctx, {"modA.bas"}, snimak_pre)
            tvrdi(sada_fajl == snimak_pre and sada_cela == snimak_pre
                  and _otisak(ctx["snimak"]) == snimak_pre,
                  "D-6: prava obnova ne vraca kopiju u stanje snimka (po fajlu "
                  "%s, cela %s), ili dira snimak" % (sada_fajl == snimak_pre,
                                                     sada_cela == snimak_pre))
        finally:
            shutil.rmtree(ctx["radni"], ignore_errors=True)
        ctx, problem = _pripremi_radni(_LazniSab(["s-zaostala"]), koren)
        shutil.rmtree(ctx["radni"], ignore_errors=True)
        tvrdi("ZATECEN SABOTIRAN" in problem and "s-zaostala" in problem,
              "D-6: snimak sa zatecenom sabotazom se ne odbija pre baze")
        tvrdi(_otisak(koren) == pre_koren, "D-6: priprema je pisala u izvor")
    finally:
        shutil.rmtree(koren, ignore_errors=True)

    # Ozicenje main-a: tekst se cita, jer main trazi Excel.
    izvor_main = inspect.getsource(main)
    tvrdi("ctx, problem = _pripremi_radni(sab, radnika=args.radnika)" in izvor_main,
          "D-6: main ne pravi snimak i kopiju pre prve baze")
    tvrdi("koren_pre = pre\n" in izvor_main
          and "_koren_menjan(pre, koren_tokom, _otisak())" in izvor_main,
          "D-6: main poredi radno stablo sa kasnijim uzorkom umesto sa snimkom")
    for ime, igla in (("_prolaz(", "baz, c)"), ("_baza_zelena(", 'src=ctx["kopija"]'),
                      ("_izolacija(", 'ctx["kopija"]')):
        redovi = [r for r in izvor_main.splitlines() if ime in r]
        tvrdi(redovi and all(igla in r for r in redovi),
              "D-6: main zove %s bez kopije -- mera bi isla nad src-vba" % ime)
    with open(os.path.abspath(__file__), encoding="utf-8") as fh:
        ceo = fh.read()
    vrati_rt = '"tools/sabotaza.py", ' + '"--vrati"'
    tvrdi(vrati_rt not in ceo,
          "D-6: dokaz.py i dalje zove --vrati nad radnim stablom")
    return nalazi


def _self_test_r4() -> list:
    """R-4 u oba smera: redosled i izolacija radnika, STOP, kapija starta,
    ponovljen start, zaostali Excel, baza izbora jednom, serijska traka."""
    nalazi = []

    def tvrdi(uslov, opis):
        if not uslov:
            nalazi.append("R-4: " + opis)

    _STOP.clear()
    # --- redosled: rezultati u redosledu poslova, iako posao 2 zavrsi PRVI ---
    gotov2 = threading.Event()
    aktivni, sudari, pozvani = set(), [], []
    brava = threading.Lock()

    def posao(k, cekaj_na=None):
        def fn(c):
            with brava:
                if c["id"] in aktivni:
                    sudari.append(c["id"])
                aktivni.add(c["id"])
                pozvani.append(k)
            if cekaj_na is not None:
                cekaj_na.wait(2.0)
            if k == 2:
                gotov2.set()
            with brava:
                aktivni.discard(c["id"])
            return (k, c["id"])
        return fn

    bazen3 = [{"id": i} for i in (1, 2, 3)]
    rez = list(_izvrsi([posao(0, gotov2), posao(1, gotov2), posao(2)], bazen3))
    tvrdi([r[0] for r in rez] == [0, 1, 2],
          "rezultati ne stizu u redosledu poslova (%r)" % [r[0] for r in rez])
    tvrdi(sorted(pozvani) == [0, 1, 2], "posao se izvrsava vise ili manje puta od jednom")
    tvrdi(not sudari and len({r[1] for r in rez}) == 3,
          "dva posla istovremeno dele radnika (%r)" % sudari)
    tvrdi(not _STOP.is_set(), "zatvaranje toka posle svih rezultata postavlja STOP")

    # Radnik se vraca TACNO jednom: vise poslova nego radnika, a posao 2 traje
    # duze od posla 1 -- dvaput vracen radnik bi tada bio kod posla 2 i 3 u isto
    # vreme.
    sudari[:] = []

    def trajanje(k, s):
        def fn(c):
            with brava:
                if c["id"] in aktivni:
                    sudari.append((k, c["id"]))
                aktivni.add(c["id"])
            time.sleep(s)
            with brava:
                aktivni.discard(c["id"])
            return k
        return fn

    rez = list(_izvrsi([trajanje(0, 0.0), trajanje(1, 0.02), trajanje(2, 0.12),
                        trajanje(3, 0.0)], [{"id": "A"}, {"id": "B"}]))
    tvrdi(rez == [0, 1, 2, 3] and not sudari,
          "radnik je kod dva posla istovremeno (%r)" % sudari)

    # Jedan radnik: isti redosled kao serijski, i dalje bez deljenja.
    pozvani[:] = []
    rez = list(_izvrsi([posao(k) for k in range(4)], [{"id": 9}]))
    tvrdi([r[0] for r in rez] == [0, 1, 2, 3] and pozvani == [0, 1, 2, 3],
          "--radnika 1 ne izvrsava poslove serijski, redom")

    # STOP pre pocetka: nijedan posao ne krece.
    pozvani[:] = []
    _STOP.set()
    rez = list(_izvrsi([posao(k) for k in range(3)], bazen3))
    _STOP.clear()
    tvrdi(rez == [None, None, None] and not pozvani,
          "posao krece posle STOP-a")

    # GRANICA: posao koji vrati STOP zaustavlja sve KASNIJE, pre nego sto
    # njegov radnik (sa nevracenom kopijom) ode sledecem poslu. Jedan radnik =
    # najgori slucaj: bez granice bi posao 2 krenuo odmah, pre potrosaca.
    pozvani[:] = []

    def sa_stopom(k):
        def fn(c):
            pozvani.append(k)
            return {"stop": "REVERT-FAIL"} if k == 1 else k
        return fn

    with contextlib.closing(_izvrsi([sa_stopom(k) for k in range(4)],
                                    [{"id": 1}])) as tok:
        rez = list(tok)
    _STOP.clear()
    tvrdi(pozvani == [0, 1] and rez[2:] == [None, None],
          "posao posle STOP rezultata krece nad nevracenom kopijom (%r)" % pozvani)

    # Dva radnika: radnik B vrati STOP odmah, a A zavrsi malo kasnije i uzme
    # SLEDECI posao -- iz reda moze da dobije bas B (nevracenu kopiju). Granica
    # mora da bude postavljena PRE nego sto B ode nazad u red.
    pokrenuti = []

    def na_radniku(k):
        def fn(c):
            pokrenuti.append((k, c["id"]))
            if k == 0:
                time.sleep(0.03)
            return {"stop": "REVERT-FAIL"} if k == 1 else k
        return fn

    with contextlib.closing(_izvrsi([na_radniku(k) for k in range(4)],
                                    [{"id": "A"}, {"id": "B"}])) as tok:
        list(tok)
    _STOP.clear()
    tvrdi(not [x for x in pokrenuti if x[0] > 1],
          "posao posle STOP-a krenuo na radniku sa nevracenom kopijom (%r)" % pokrenuti)

    # Izuzetak u petlji potrosaca: closing zatvara tok, STOP se postavlja, a
    # poslovi koji nisu krenuli ne krecu (ranije: svi, satima, bez izlaza).
    pozvani[:] = []
    kreni = threading.Event()

    def spori(k):
        def fn(c):
            pozvani.append(k)
            if k == 1:
                kreni.wait(2.0)
            return k
        return fn

    try:
        with contextlib.closing(_izvrsi([spori(k) for k in range(4)],
                                        [{"id": 1}])) as tok:
            for r in tok:
                kreni.set()
                raise RuntimeError("potrosac pukao")
    except RuntimeError:
        pass
    stop_postavljen = _STOP.is_set()
    _STOP.clear()
    tvrdi(stop_postavljen and 2 not in pozvani and 3 not in pozvani,
          "izuzetak potrosaca ostavlja bazen da pali poslove (%r)" % pozvani)

    # --- kapija starta -------------------------------------------------------
    sat, spavanja = [100.0], []

    def spavaj(s):
        spavanja.append(round(s, 3))
        sat[0] += s

    stari = dict(_START)
    try:
        _START.update(razmak=1.0, poslednji=0.0)
        _kapija_starta(lambda: sat[0], spavaj)
        _kapija_starta(lambda: sat[0], spavaj)
        sat[0] += 0.4
        _kapija_starta(lambda: sat[0], spavaj)
    finally:
        _START.update({k: v for k, v in stari.items() if k != "lock"})
    tvrdi(spavanja == [1.0, 0.6], "kapija starta ne drzi razmak (%r)" % spavanja)

    # --- ponovljen start: samo kad Excel ni ne krene -------------------------
    PRE_SUITE = "FATAL   Excel se nije pokrenuo\n"
    U_SUITE = "SUITE   FAIL   RunAllTests (3.0s)\nFATAL   ubijen u fazi\n"
    OK = "TESTS   RunAllTests: 1 ukupno, 0 palo\n"
    for opis, odgovori, ocekivano_poziva, stop in (
            ("FATAL pre suite-a se ponavlja jednom", [PRE_SUITE, OK], 2, False),
            ("FATAL posle starta suite-a se ne ponavlja", [U_SUITE, OK], 1, False),
            ("dvaput FATAL pre suite-a: drugi se prijavljuje", [PRE_SUITE, PRE_SUITE], 2, False),
            ("posle STOP-a run uopste ne krece", [PRE_SUITE, OK], 0, True),
            ("otkazana baza se ne ponavlja posle ubijanja", [PRE_SUITE, OK], 1, "otkaz")):
        red_odg = list(odgovori)
        broj = [0]

        def lazni(*a, **kw):
            broj[0] += 1
            return _LazniOdgovor(0, red_odg.pop(0))

        pravi = _pusti
        globals()["_pusti"] = lazni
        otkaz = threading.Event()
        if stop is True:
            _STOP.set()

        def lazni_otkaz(*a, **kw):
            otkaz.set()             # otkaz stize DOK prvi run traje: Excel ubijen
            broj[0] += 1
            return _LazniOdgovor(2, PRE_SUITE)

        if stop == "otkaz":
            globals()["_pusti"] = lazni_otkaz
        pon_pre = _START["ponovljeno"]
        try:
            _pusti_suite("x", "tools/run_vba.py",
                         prekini=otkaz if stop == "otkaz" else None)
        finally:
            globals()["_pusti"] = pravi
            _STOP.clear()
        pon = _START["ponovljeno"] - pon_pre
        _START["ponovljeno"] = pon_pre
        tvrdi(broj[0] == ocekivano_poziva and (stop != "otkaz" or pon == 0),
              "%s (poziva %d, ponovljenih %d)" % (opis, broj[0], pon))

    # --- zaostali Excel: samo uz isti PID I isto vreme nastanka --------------
    izl = tempfile.mkdtemp(prefix="dokaz_r4_")
    try:
        for opis, vreme_zivog, ocekivano in (("isto vreme", "T1", True),
                                             ("drugo vreme (PID ponovo dat)", "T2", False)):
            with open(os.path.join(izl, "excel.pid"), "w", encoding="ascii") as fh:
                json.dump({"pid": 77, "kreiran": "T1"}, fh)
            ubijeni = []
            dobijeno = _ubij_zaostali(izl, otvori=lambda pid: ("h", pid),
                                      vreme=lambda h: vreme_zivog,
                                      ubij=lambda h: ubijeni.append(h))
            tvrdi(dobijeno == ocekivano and bool(ubijeni) == ocekivano,
                  "zaostali Excel (%s): ubijen=%r" % (opis, ubijeni))
        try:
            os.remove(os.path.join(izl, "excel.pid"))
        except OSError:
            pass
        tvrdi(not _ubij_zaostali(izl, otvori=lambda pid: "h", vreme=lambda h: "T1",
                                 ubij=lambda h: nalazi.append("R-4: ubija bez excel.pid")),
              "bez excel.pid se nesto ubija")
    finally:
        shutil.rmtree(izl, ignore_errors=True)

    # --- baza izbora: jednom po kljucu i kad je traze dva radnika odjednom ---
    poziva_baze, pusti_bazu = [0], threading.Event()
    pravi_bz = _baza_zelena

    def lazna_baza(suite, testovi=None, src=None, izlaz=None, prekini=None):
        poziva_baze[0] += 1
        pusti_bazu.wait(2.0)
        return True, "ok"

    globals()["_baza_zelena"] = lazna_baza
    baze, ishodi = {}, []
    try:
        niti = [threading.Thread(target=lambda: ishodi.append(
            _baza_izbora(baze, ("S", ("T",)), "S", ["T"], {"kopija": "k", "izlaz": "i"})))
            for _ in range(2)]
        for n in niti:
            n.start()
        time.sleep(0.02)
        pusti_bazu.set()
        for n in niti:
            n.join(3.0)
    finally:
        globals()["_baza_zelena"] = pravi_bz
    tvrdi(poziva_baze[0] == 1 and ishodi == [(True, "ok"), (True, "ok")],
          "ista baza izbora se meri vise puta (%d) ili ne stize svima" % poziva_baze[0])

    # --- PREKID: run u toku kad je STOP postavljen ne daje presudu -----------
    def prolaz_uz_stop():
        pozivi = []

        def lazni_pusti(*a, **kw):
            pozivi.append(a[1:])
            if a[1].endswith("run_vba.py"):
                _STOP.set()
                return _LazniOdgovor(0, "FAIL T_A -- razlog a\n")
            return _LazniOdgovor(0)

        pravi = _pusti, _otisak, _obnovi
        globals()["_pusti"] = lazni_pusti
        globals()["_otisak"] = lambda put=None: "P"
        globals()["_obnovi"] = lambda c, f=None: None
        try:
            return _prolaz([_s("a", tvrdnja="razlog a")], "P", False, ctx=_LAZNI_CTX)
        finally:
            globals()["_pusti"], globals()["_otisak"], globals()["_obnovi"] = pravi
            _STOP.clear()

    rez = prolaz_uz_stop()
    tvrdi(rez["stop"] == "PREKID"
          and [(st, p) for _s2, p, st, _c in rez["ishodi"]] == [("PREKID", False)],
          "run prekinut STOP-om se ocenjuje kao ishod sabotaze (%r)" % rez["stop"])

    # --- serijska traka i bazen ---------------------------------------------
    sef = [_s("x", suite="RunSEFTestSuite")]
    tvrdi(_serijska_traka(sef, 1) == "" and "RunSEFTestSuite" in _serijska_traka(sef, 3)
          and _serijska_traka([_s("a"), _s("b", suite=SUITE_BFP)], 5) == "",
          "serijska traka pusti suite van PARALELNE_SUITE u radnike, ili odbija dozvoljene")
    tvrdi("dele" in _proveri_bazen([dict(_LAZNI_CTX, izlaz="/i1"),
                                    dict(_LAZNI_CTX, izlaz="/i2")], _LAZNI_CTX)
          and "dele" in _proveri_bazen([dict(_LAZNI_CTX, kopija="/k1", izlaz="/i"),
                                        dict(_LAZNI_CTX, kopija="/k2", izlaz="/i")],
                                       _LAZNI_CTX)
          and _proveri_bazen([dict(_LAZNI_CTX, kopija="/k1", izlaz="/i1"),
                              dict(_LAZNI_CTX, kopija="/k2", izlaz="/i2")],
                             _LAZNI_CTX) == "",
          "bazen sa deljenom kopijom ili izlazom se prihvata (ili ispravan odbija)")

    # --- istorija radnika: spaja se jednom, a zauzet tests/ je ne gubi --------
    rv = _run_vba_modul()
    koren = tempfile.mkdtemp(prefix="dokaz_ist_")
    try:
        cilj = os.path.join(koren, "tests", "run_vba_istorija.jsonl")
        os.makedirs(os.path.dirname(cilj))
        bz = []
        for k in (1, 2):
            izl = os.path.join(koren, "r%d" % k)
            os.makedirs(izl)
            with open(os.path.join(izl, "run_vba_istorija.jsonl"), "w") as fh:
                fh.write(json.dumps({"rc": 0, "k": k}) + "\n")
            bz.append({"izlaz": izl})
        prvi = _spoji_istoriju(bz, 2, cilj)
        drugi = _spoji_istoriju(bz, 2, cilj)
        with open(cilj) as fh:
            redovi = [json.loads(r) for r in fh if r.strip()]
        tvrdi(prvi == "" and drugi == "" and len(redovi) == 2
              and all(r.get("dokaz_radnika") == 2 for r in redovi),
              "istorija radnika se ne spaja tacno jednom (%d redova)" % len(redovi))
        with open(os.path.join(bz[0]["izlaz"], "run_vba_istorija.jsonl"), "w") as fh:
            fh.write(json.dumps({"rc": 0, "k": 3}) + "\n")
        drzi = rv._zakljucaj(os.path.join(os.path.dirname(cilj), ".run_vba.lock"))
        pravi_zak = rv._zakljucaj
        rv._zakljucaj = lambda put, cekaj=0.0: pravi_zak(put, 0.0)
        try:
            zauzet = _spoji_istoriju(bz, 2, cilj)
        finally:
            rv._zakljucaj = pravi_zak
            rv._pusti_bravu(drzi)
        rezerve = [f for f in os.listdir(os.path.dirname(cilj))
                                  if f.startswith("run_vba_istorija.dokaz_")]
        tvrdi(zauzet.startswith("istorija je u ") and len(rezerve) == 1,
              "zauzet tests/ gubi istoriju radnika (%r)" % zauzet)
    finally:
        shutil.rmtree(koren, ignore_errors=True)

    # --- baze unapred: redom, na rezervisanom radniku; otkaz ne ceka suite ---
    redosled = []
    pravi_bz = _baza_zelena

    def lazna(suite, testovi=None, src=None, izlaz=None, prekini=None):
        redosled.append((suite, src, izlaz))
        return True, "TESTS %s: 1 ukupno, 0 palo" % suite

    globals()["_baza_zelena"] = lazna
    try:
        bu = _BazeUnapred(["S1", "S2"], {"kopija": "/k5", "izlaz": "/i5"},
                          ubij=lambda: None)
        dobijeno = bu.sacekaj()
    finally:
        globals()["_baza_zelena"] = pravi_bz
    tvrdi(redosled == [("S1", "/k5", "/i5"), ("S2", "/k5", "/i5")]
          and sorted(dobijeno) == ["S1", "S2"] and bu._ex not in _AKTIVNI,
          "baze unapred se ne mere redom na rezervisanom radniku (%r), ili "
          "izvrsilac ostaje prijavljen" % redosled)

    redosled[:] = []
    ubijeno, pusti = [], threading.Event()

    def visi(suite, testovi=None, src=None, izlaz=None, prekini=None):
        tvrdi(prekini is not None, "baza unapred ne predaje svoj otkaz pozivu")
        redosled.append(suite)
        pusti.wait(2.0)
        return False, "ubijen"

    globals()["_baza_zelena"] = visi
    try:
        bu = _BazeUnapred(["S1", "S2"], {"kopija": "/k5", "izlaz": "/i5"},
                          ubij=lambda: (ubijeno.append(1), pusti.set()))
        time.sleep(0.02)
        u_registru = bu in _UNAPRED
        t0 = time.time()
        bu.otkazi()
        trajalo = time.time() - t0
    finally:
        globals()["_baza_zelena"] = pravi_bz
    tvrdi(u_registru and bu not in _UNAPRED,
          "baza unapred nije u registru dok radi (prekid je ne bi otkazao), "
          "ili ostaje u njemu posle otkaza")
    tvrdi(redosled == ["S1"] and ubijeno and trajalo < 1.5,
          "otkaz baze unapred ne prekida tekucu suite ili pusta sledecu (%r, %.2fs)"
          % (redosled, trajalo))

    # Faza 3 treba samo S1: cim je S1 gotova, S2 (puna suite) se otkazuje.
    redosled[:] = []
    ubijeno2, pusti2 = [], threading.Event()

    def s1_brza(suite, testovi=None, src=None, izlaz=None, prekini=None):
        redosled.append(suite)
        if suite == "S2":
            pusti2.wait(2.0)
            return False, "ubijen"
        return True, "ok S1"

    globals()["_baza_zelena"] = s1_brza
    try:
        bu = _BazeUnapred(["S1", "S2"], {"kopija": "/k5", "izlaz": "/i5"},
                          ubij=lambda: (ubijeno2.append(1), pusti2.set()))
        t0 = time.time()
        dobijeno = bu.sacekaj({"S1"})
        trajalo = time.time() - t0
    finally:
        globals()["_baza_zelena"] = pravi_bz
    tvrdi(dobijeno.get("S1") == (True, "ok S1") and ubijeno2 and trajalo < 1.5
          and bu not in _UNAPRED,
          "faza 3 ceka baze koje joj ne trebaju (%.2fs, ubijeno=%r)" % (trajalo, ubijeno2))

    # --- ozicenje main-a i _prolaz-a (tekst; main trazi Excel) -----------------
    izvor_main = inspect.getsource(main)
    izvor_prolaz = inspect.getsource(_prolaz)
    tvrdi(izvor_main.count("_izvrsi(") >= 3
          and "_serijska_traka(stavke, args.radnika)" in izvor_main
          and "_zavrsi_istoriju(bazen, args.radnika)" in izvor_main,
          "main ne pusta sve tri faze kroz radnike, ili preskace serijsku traku/istoriju")
    tvrdi(all('izlaz=ctx["izlaz"]' in r for r in izvor_main.splitlines()
              if "_baza_zelena(" in r),
          "baza cele suite ne ide u izlaz radnika (pisala bi u tests/)")
    tvrdi("    if args.radnika > 1 and izbor:\n" in izvor_main
          and "unapred = _BazeUnapred(moguce, bazen[-1])" in izvor_main
          and "return bazen[:-1]" in izvor_main
          and izvor_main.count("bazen_faze())) as tok:") == 3
          and 0 <= izvor_main.find("_izvrsi([posao([s], False) for s in sve], bazen_faze())")
          < izvor_main.find("unapred.sacekaj(potrebne_3)")
          and "unapred.otkazi()" in izvor_main,
          "baze unapred: pogresan uslov, rezervisani radnik nije izuzet iz faza, "
          "ili faza 3 ne ceka (ne otkazuje) baze unapred")
    # Rana mutacija: ocena TEK uz zelenu bazu, a STOP ne ceka punu suite baze.
    rana = izvor_main[izvor_main.find("# RANA MUTACIJA"):]
    tvrdi(izvor_main.count('upisi(ishod["ishodi"])') == 3
          and "                    ok, opis = cele_baze.get(s[\"suite\"],\n" in rana
          and "                    if ok:\n                        upisi(ishod[\"ishodi\"])" in rana
          and "if unapred is not None and (not ishod[\"stop\"] or unapred.gotovo()):" in rana,
          "rana mutacija: priznaje se bez zelene baze cele suite, ili STOP ceka "
          "punu suite baze")
    tvrdi(izvor_main.count("contextlib.closing(_izvrsi(") >= 3
          and "tok.close()" not in izvor_main,
          "faze ne zatvaraju tok kroz contextlib.closing (izuzetak bi ostavio bazen)")
    tvrdi("f.result(timeout=" in inspect.getsource(_izvrsi),
          "glavna nit ceka rezultat bez koraka (Ctrl+C na Windows-u ne stize)")
    with open(os.path.abspath(__file__), encoding="utf-8") as fh:
        ceo = fh.read()
    glavni = ceo[ceo.rfind('if __name__ == "__main__":'):]
    tvrdi(0 <= glavni.find("ex.shutdown(wait=True, cancel_futures=True)")
          < glavni.find("shutil.rmtree(_RADNI"),
          "Ctrl+C brise radni folder pre nego sto radnici stanu")
    tvrdi(0 <= glavni.find("for bu in list(_UNAPRED):\n                bu.otkazi()")
          < glavni.find("ex.shutdown(wait=True, cancel_futures=True)")
          and "    except BaseException:\n" in glavni
          and glavni.count("bu.otkazi()") == 2,
          "prekid ili izuzetak ne otkazuju bazu unapred (cekali bi punu suite)")
    tvrdi("_baza_izbora(baze, kljuc, suite, testovi, ctx)" in izvor_prolaz
          and izvor_prolaz.count("_pusti_suite(") == 1
          and 'izlaz=ctx.get("izlaz")' in izvor_prolaz,
          "_prolaz ne ide kroz zajednicku bazu izbora ili kapiju starta")
    return nalazi


def _self_test(tiho: bool = False) -> int:
    nalazi = []
    _STOP.clear()
    nalazi.extend(_self_test_r4())

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
            return {"stop": "", "ishodi": [], "ostatak": [], "cela_suite": []}, []

    def suita(pozivi):
        return [p for p in pozivi if p and p[0].endswith("run_vba.py")]

    rez, pozivi = prolaz(g3, SVE)
    tvrdi(len(suita(pozivi)) == 1,
          "PROLAZ: grupa od tri clana ne pusta suite TACNO jednom")
    tvrdi(len(rez["ishodi"]) == 3 and all(p for _s2, p, _st, _c in rez["ishodi"])
          and not rez["ostatak"],
          "PROLAZ: grupa u kojoj su sve tvrdnje pale ne priznaje sve clanove")
    tvrdi(pozivi and pozivi[-1][0] == "<obnova>",
          "PROLAZ: kopija se ne vraca iz snimka posle prolaza")

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

    rez, pozivi = prolaz(g3, SVE, potpis="DRUGI")
    tvrdi(rez["stop"].startswith("REVERT-FAIL")
          and not any(p for _s2, p, _st, _c in rez["ishodi"]),
          "PROLAZ: nevracen izvor ne zaustavlja dokaz niti gasi priznanja")
    tvrdi(("<obnova>", "<cela>") in pozivi,
          "PROLAZ: kopija koja se ne vraca po fajlu ne obnavlja se CELA pre STOP-a")

    # --- KI-036: prolaz nad mutiranim izvorom ne pise marker zelenog ---------
    _rez, pozivi = prolaz([g3[0]], SVE, grupno=False)
    tvrdi([("--bez-markera" in r) for r in suita(pozivi)] == [True],
          "KI-036: prolaz nad mutiranim izvorom moze da upise marker zelenog")
    tvrdi("--bez-markera" not in _run_suite(SUITE_ALL)
          and "--bez-markera" not in _run_suite(SUITE_ALL, ["T_A"]),
          "KI-036: bazni prolaz (cist izvor) ne sme da gasi upis markera")

    # --- FATAL iz run_vba: ubijen prolaz nije ishod sabotaze -----------------
    UBIJEN = ("FATAL   Excel se nije vratio 15s posle dijaloga u fazi "
              "'RunAllTests' -- proces je ubijen.\n")
    rez, _p = prolaz([g3[0]], UBIJEN, grupno=False)
    tvrdi([(s["ime"], p, st.split(" --")[0]) for s, p, st, _c in rez["ishodi"]]
          == [("a", False, "EXCEL UBIJEN")],
          "FATAL: ubijen prolaz se ocenjuje kao ishod sabotaze (NE OBARA NISTA)")
    rez, _p = prolaz(g3, UBIJEN)
    tvrdi(sorted(s["ime"] for s, _r in rez["ostatak"]) == ["a", "b", "c"]
          and not rez["ishodi"],
          "FATAL: ubijen grupni prolaz se ne rastura na pojedinacna merenja")
    pravi_pusti = _pusti
    globals()["_pusti"] = lambda *a, **kw: _LazniOdgovor(2, UBIJEN)
    try:
        ok, opis = _baza_zelena(SUITE_ALL)
    finally:
        globals()["_pusti"] = pravi_pusti
    tvrdi(not ok and "EXCEL UBIJEN" in opis,
          "FATAL: ubijena baza se prijavljuje kao 'nema reda TESTS' umesto razloga")

    # --- izbor: samo ciljani testovi, uz bazu nad ISTIM izborom -------------
    #
    # Baza je "TESTS ... 0 palo" nad cistim izvorom; lazni runner je razlikuje
    # od prolaza sa mutacijom po tome da li je neka mutacija primenjena.
    IZB = frozenset({SUITE_ALL})
    ZELENA = "TESTS   RunAllTests: 3 ukupno, 0 palo\n"
    CRVENA = "TESTS   RunAllTests: 3 ukupno, 1 palo\nFAIL T_B -- sam ne prolazi\n"

    def testovi_poziva(p):
        return sorted(p[i + 1] for i in range(len(p) - 1) if p[i] == "--test")

    def je_primena(p):
        return p and p[0].endswith("sabotaza.py")

    rez, pozivi = prolaz(g3, SVE, izbor=IZB, baze={}, izlaz_baze=ZELENA)
    runovi = suita(pozivi)
    tvrdi(len(runovi) == 2 and all(testovi_poziva(r) == ["T_A", "T_B", "T_C"]
                                   for r in runovi),
          "IZBOR: baza i prolaz ne pustaju TACNO testove clanova")
    prva_primena = next((i for i, p in enumerate(pozivi) if je_primena(p)), -1)
    tvrdi(pozivi and pozivi[0][0].endswith("run_vba.py")
          and 0 < prva_primena,
          "IZBOR: baza izbora se ne meri PRE prve mutacije")
    tvrdi(len(rez["ishodi"]) == 3 and all(p for _s2, p, _st, _c in rez["ishodi"]),
          "IZBOR: zelena baza + pad svih tvrdnji ne priznaje clanove")

    baze = {}
    prolaz(g3, SVE, izbor=IZB, baze=baze, izlaz_baze=ZELENA)
    _rez, pozivi = prolaz(g3, SVE, izbor=IZB, baze=baze, izlaz_baze=ZELENA)
    tvrdi(len(suita(pozivi)) == 1,
          "IZBOR: baza istog izbora se placa ponovo umesto da se zapamti")

    rez, pozivi = prolaz(g3, SVE, izbor=IZB, baze={}, izlaz_baze=CRVENA)
    tvrdi(not any(je_primena(p) for p in pozivi) and len(suita(pozivi)) == 1,
          "IZBOR: crvena baza grupe ipak primenjuje mutacije ili pusta suite")
    tvrdi(sorted(s["ime"] for s, _r in rez["ostatak"]) == ["a", "b", "c"]
          and not rez["ishodi"] and not rez["cela_suite"],
          "IZBOR: grupa sa crvenom bazom se ne rastura na pojedinacna merenja")

    rez, pozivi = prolaz([g3[1]], SVE, grupno=False, izbor=IZB, baze={},
                         izlaz_baze=CRVENA)
    tvrdi([s["ime"] for s, _r in rez["cela_suite"]] == ["b"]
          and not rez["ishodi"] and not any(je_primena(p) for p in pozivi),
          "IZBOR: test koji SAM ne prolazi se ne salje na merenje celom suite-om")

    rez, _p = prolaz([g3[0]], SVE, grupno=False, izbor=IZB, baze={},
                     izlaz_baze="TESTS   RunAllTests: 0 ukupno, 0 palo\n")
    tvrdi([s["ime"] for s, _r in rez["cela_suite"]] == ["a"],
          "IZBOR: izbor bez ijedne provere se prihvata kao zelena baza")

    # Izbor koji NE obori tvrdnju nije presuda: clan ide na celu suite. Bez
    # izbora isti ishod ostaje presuda -- povratak ne sme da vazi i tamo.
    NISTA = "TESTS   RunAllTests: 3 ukupno, 0 palo\n"
    rez, pozivi = prolaz([g3[0]], NISTA, grupno=False, izbor=IZB, baze={},
                         izlaz_baze=ZELENA)
    tvrdi([s["ime"] for s, _r in rez["cela_suite"]] == ["a"] and not rez["ishodi"]
          and rez["cela_suite"][0][1] == "u izboru: NE OBARA NISTA",
          "IZBOR: sabotaza koju izbor ne obori dobija presudu bez cele suite")
    tvrdi(rez.get("izbor") is True and len(suita(pozivi)) == 2,
          "IZBOR: prolaz izborom nije oznacen kao izbor (verdikt bi bio golo DOKAZANO)")
    rez, _p = prolaz([g3[0]], NISTA, grupno=False)
    tvrdi(rez.get("izbor") is False and not rez["cela_suite"]
          and [(s["ime"], p, st) for s, p, st, _c in rez["ishodi"]]
          == [("a", False, "NE OBARA NISTA")],
          "IZBOR: povratak na celu suite vazi i za prolaz BEZ izbora")
    rez, _p = prolaz(g3, NISTA, izbor=IZB, baze={}, izlaz_baze=ZELENA)
    tvrdi(sorted(s["ime"] for s, _r in rez["ostatak"]) == ["a", "b", "c"]
          and not rez["cela_suite"],
          "IZBOR: clan grupe koji ne obori tvrdnju preskace pojedinacno merenje")

    # --- D-6: mutacija ide u kopiju, src-vba se ne pise ----------------------
    nalazi.extend(_self_test_d6(prolaz, suita))

    # --- verdikt: svaki rezim koji meri manje nosi svoje ime ----------------
    for ocekivano, args in (("DOKAZANO", (True, 0, 0)),
                            ("DOKAZANO (grupno)", (True, 3, 0)),
                            ("DOKAZANO (izbor)", (True, 0, 2)),
                            ("DOKAZANO (grupno, izbor)", (True, 3, 2)),
                            ("DOKAZANO (izbor, snimak abc)", (True, 0, 2, "abc")),
                            ("NIJE DOKAZANO", (False, 3, 2, "abc")),
                            ("NIJE DOKAZANO", (False, 3, 2))):
        tvrdi(_verdikt(*args) == ocekivano,
              "VERDIKT: %r daje %r umesto %r" % (args, _verdikt(*args), ocekivano))
    # main() mora da broji priznanja izbora i da ih preda verdiktu. Tekst main-a
    # se cita, jer main trazi Excel; igla je samo u main-u, ne i ovde.
    izvor_main = inspect.getsource(main)
    tvrdi("_verdikt(ok, grupno_priznatih, izborom_priznatih," in izvor_main
          and "pre if koren_menjan else" in izvor_main
          and izvor_main.count("izborom_priznatih += ") == 2,
          "VERDIKT: main ne broji priznanja izbora, ili ne imenuje snimak kad se "
          "radno stablo menjalo")

    # --- izolacija: svaki razlicit ciljani test SAM, bez mutacija -----------
    izol_pozivi = []

    def izol_pusti(*a, **kw):
        izol_pozivi.append(tuple(a[1:]))
        crven = "T_B" in a
        return _LazniOdgovor(0, "TESTS   RunAllTests: 1 ukupno, %d palo\n"
                             % (1 if crven else 0))

    stavke_izol = [_s("a"), _s("a2"), _s("b", test="T_B", kljuc="T_B"),
                   _s("k", test="T21_X", kljuc="T21", suite=SUITE_BANKA)]
    pravi_pusti = _pusti
    globals()["_pusti"] = izol_pusti
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            crveni = _izolacija(stavke_izol, {SUITE_ALL})
    except Exception as e:                  # pad je nalaz, ne goli traceback
        crveni = None
        nalazi.append("IZOLACIJA: pukla -- %r" % (e,))
    finally:
        globals()["_pusti"] = pravi_pusti
    tvrdi(crveni is not None and [c[1] for c in crveni] == ["T_B"],
          "IZOLACIJA: ne vraca bas test koji sam ne prolazi")
    tvrdi(sorted(tuple(p[p.index("--test") + 1:]) for p in izol_pozivi
                 if "--test" in p) == [("T_A",), ("T_B",)]
          and len(izol_pozivi) == 2,
          "IZOLACIJA: test se pusta vise puta, sa partnerom, ili i suite bez izbora")

    banka = [_s("k", test="T21_X", kljuc="T21", tvrdnja="t", suite=SUITE_BANKA)]
    _rez, pozivi = prolaz(banka, "PAO   T21 t\n", grupno=False, izbor=IZB,
                          baze={})
    tvrdi(len(suita(pozivi)) == 1 and not testovi_poziva(suita(pozivi)[0]),
          "IZBOR: suite bez ulazne tacke izbora dobija --test ili bazu izbora")

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
        # D-6: nema sta da se vraca -- mutacije su isle u kopiju, src-vba nije
        # pisan. Ranije je ovde stajao `sabotaza.py --vrati` nad radnim stablom,
        # koji bi usput "vratio" i rucno primenjenu sabotazu operatera.
        # R-4: Excel radnika koji je preziveo svoj run_vba ubija se po excel.pid
        # (samo uz isto vreme nastanka), a radni folder se brise.
        _STOP.set()
        # PRVO sacekaj radnike: brisanje radnog foldera dok run_vba jos pise u
        # njega ostavilo bi polu-obrisan folder (i lazno "obrisan"). Baza
        # unapred je puna suite u jednom COM pozivu -- ne ceka se nego OTKAZUJE.
        try:
            for bu in list(_UNAPRED):
                bu.otkazi()
            for ex in list(_AKTIVNI):
                ex.shutdown(wait=True, cancel_futures=True)
        except KeyboardInterrupt:
            print("\nprekinuto dvaput -- radnici se jos gase; src-vba nije diran; "
                  "radni folder ostaje: %s" % _RADNI, file=sys.stderr)
            sys.exit(130)
        ubijeno = sum(1 for c in _BAZEN if _ubij_zaostali(c.get("izlaz")))
        sacuvana = _zavrsi_istoriju(_BAZEN, len(_BAZEN)) if _BAZEN else True
        if _RADNI and sacuvana:
            shutil.rmtree(_RADNI, ignore_errors=True)
        print("\nprekinuto -- src-vba nije diran (dokaz radi nad kopijom); "
              "radni folder %s%s"
              % ("obrisan" if sacuvana else "ostaje: %s" % _RADNI,
                 ", zaostalih Excel-a ubijeno: %d" % ubijeno if ubijeno else ""),
              file=sys.stderr)
        sys.exit(130)
    except BaseException:
        # Neuhvacen izuzetak iz main-a: baza unapred ne sme da drzi proces dok
        # se puna suite ne zavrsi (atexit bi cekao njenu nit).
        _STOP.set()
        for bu in list(_UNAPRED):
            bu.otkazi()
        raise
