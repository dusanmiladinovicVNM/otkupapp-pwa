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
    python tools/dokaz.py --prijem --radnika 4 # G1: izbor prema celoj suiti, po stavci

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

6. PRIJEM (`--prijem [N]`, G1) MERI KOLIKO IZBOR TVRDI MANJE. Nad uzorkom
   (ciljani sloj -- S-4a, testovi mehanizma izbora, POZNATI_NALAZI_DOKAZ -- i N
   slucajnih po deterministickom pravilu, v. `_g1_uzorak`) svaka stavka ide
   izborom, verdikt izbora se zapamti (`_prijem_ishod`, u `najavi`, PRE
   knjizenja), pa se ista stavka meri celom suite-om (faza 3, nepromenjena).
   Na kraju se dva verdikta porede po stavci: OPASNA (izbor priznaje, cela
   suite ne) ili STROZA obara prijem (rc 1); stavka koju izbor nije presudio
   sam (SKUPA) ne ulazi u imenilac, jer bi tu cela suite poredila samu sebe.
   Premalo presudjenih, STOP, NEJASNO stavka ili filter daju rc 2. Bez
   `--prijem` se nista ne menja.

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

# --- G1: prijem izbora (v. tacka 6 u docstring-u i `--prijem`) --------------
#
# G1_N slucajnih stavki ima rezervu za povratke: stavka koju izbor ne presudi
# sam (SKUPA) ne ulazi u imenilac, a ocekivani udeo povrataka nije meren.
# G1_MIN presudjenih izborom daje gornju granicu ~5,8% (95%) uz 0 razlika;
# G1_MIN_SUITE drzi da ni RAT ni BFP ne budu prazni u imeniocu. Release (O-1 V2,
# >= 100 presudjenih) je `--prijem 115`, a red G1 PRIJEM ispisuje stvarno M.
G1_N = 60
G1_MIN = 50
G1_MIN_SUITE = 15
G1_SEME = "G1"
# Testovi S-4a (fa53eb32): prolazili su SAMO u punoj suite-i, pa je njihov
# povratak na celu suite ono sto G1 mora da vidi.
G1_S4A_TESTOVI = (
    "T_Analiza_EkranUIzradi", "T_Doplata_PonavljanjeTraziPotvrdu",
    "T_Doplata_SaldoOMPoStanici", "T_AmbPartneri_ListaIzvestaja",
    "T_AmbBilans_ListaIzvestaja", "T_Gotovina_NajviseNaIzboru",
    "T_S6K4_UvidKazeOdbijanje", "Test_S6K7_LanacBezKupcaNePise",
    "T_CiljZbirna_NePoPrvomRedu", "T_Prijemnica_PomeraSamoAktivan",
    "T_IstiBrojRazliciteGeneracije_NijeIstiDokument",
    "T_AmbPoslovno_OtkupniListKaoKartica",
    "Test_PrijemnicaMissingZbirnaDoesNotAppend")
# Testovi mehanizma izbora i sabotaze dodate posle #479 -- merene samo izborom.
G1_POSLE_IZBORA_TESTOVI = (
    "T_TestIzbor_TacnoImeIJednokratno", "Test_BFP_IzborBiraCeloIme",
    "T_Journal_TihiRezimGasiZakazivanje", "Test_BFP_TihiRezimDnevnika",
    "Test_OTP_AutoKvarGrupeNeObaraOstale",
    "T_Backup_NeObaraStartINeRasteBezGranice")
# Stanja koja ne kazu nista o tvrdnji (prolaz nije stigao do kraja): razlika
# uz njih je NEJASNO, ne OPASNA.
_G1_INFRA = ("TIMEOUT", "EXCEL UBIJEN", "PREKID", "REVERT-FAIL", "STRAZA")
_G1_KRATKO = {SUITE_ALL: "RAT", SUITE_BFP: "BFP", SUITE_BANKA: "banka"}

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
    radnici = bazen if isinstance(bazen, _Radnici) else _Radnici(bazen)
    granica = [len(poslovi)]
    brava = threading.Lock()

    def omot(i, fn):
        if _STOP.is_set() or i > granica[0]:
            return None
        c = radnici.uzmi(odustani=lambda: _STOP.is_set() or i > granica[0], kljuc=i)
        if c is None:
            return None
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
            radnici.vrati(c)

    ex = concurrent.futures.ThreadPoolExecutor(max_workers=len(radnici))
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
    # Kopija mora biti u stanju snimka PRE prolaza, ne samo posle prethodnog:
    # radnik dolazi iz zajednickog bazena, a prolaz koji javi REVERT-FAIL ga
    # vraca sa mutacijom u kopiji. Mera nad njom bi bila tudja (R-4).
    zatecen = _otisak(ctx["kopija"])
    if zatecen != pre:
        rez["stop"] = ("STRAZA: kopija radnika r%s nije u stanju snimka pre "
                       "prolaza (potpis %s) -- mera bi isla nad tudjom mutacijom"
                       % (ctx.get("id"), zatecen))
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


class _Radnici:
    """Radnici (ctx) ZAJEDNICKI za sve faze i za baze unapred (R-4).

    Ko uzme radnika drzi ga do vracanja -- u jednom trenutku je radnik kod
    najvise jednog posla. Prioritetan zahtev (baza cele suite) dobija PRVOG
    slobodnog, pre poslova koji vec cekaju: inace bi baza cekala kraj cele
    faze, a ona je duza od svakog posla izbora.

    Obicni zahtevi se uslugaju po KLJUCU (indeks posla), ne po tome ko stigne:
    nit koja vrati radnika ne sme da ga odmah uzme za svoj KASNIJI posao ispred
    ranijeg koji ceka. Granica STOP-a u _izvrsi stiti samo kasnije poslove --
    raniji posao koji krene POSLE REVERT-FAIL-a kasnijeg dobio bi njegovu
    nevracenu kopiju."""

    def __init__(self, bazen: list, korak: float = 0.5):
        self.bazen = list(bazen)
        self._slobodni = list(bazen)
        self._uslov = threading.Condition()
        self._prioritetnih = 0
        self._red = []                  # obicni zahtevi koji cekaju: (kljuc, broj)
        self._broj = 0
        self._korak = korak             # koliko cesto ceka pita odustani()

    def __len__(self) -> int:
        return len(self.bazen)

    def uzmi(self, prioritet: bool = False, odustani=None, kljuc=None):
        """Slobodan radnik, ili None ako odustani() postane tacno dok se ceka.
        Obican zahtev ceka svoj red: najmanji kljuc (bez kljuca -- redosled
        dolaska) i nijedan prioritetan ispred sebe."""
        with self._uslov:
            if prioritet:
                self._prioritetnih += 1
            else:
                self._broj += 1
                moj = (self._broj if kljuc is None else kljuc, self._broj)
                self._red.append(moj)
            try:
                while not self._slobodni or (not prioritet and (
                        self._prioritetnih or min(self._red) != moj)):
                    if odustani is not None and odustani():
                        return None
                    self._uslov.wait(self._korak)
                return self._slobodni.pop(0)
            finally:
                if prioritet:
                    self._prioritetnih -= 1
                else:
                    self._red.remove(moj)
                self._uslov.notify_all()

    def vrati(self, c: dict) -> None:
        with self._uslov:
            self._slobodni.append(c)
            self._uslov.notify_all()


class _BazeUnapred:
    """Pune baze cele suite UNAPRED, NA ZAHTEV (R-4).

    Uz --izbor se puna baza trazi samo za sabotaze koje izbor ne obori (faza
    3). Bez ovoga je na kraju dokaza isla puna suite za bazu, pa puna suite za
    mutaciju, jedna za drugom. Baza zato krece RANIJE -- ali tek kad se pojavi
    prva stavka koja vodi u fazu 3 (trazi), ne spekulativno na pocetku: rez bez
    faze 3 inace placa punu suite i radnika za nju, a nista ih ne koristi.

    Radnika POZAJMLJUJE iz zajednickog bazena (prioritet) i vraca ga posle
    svake baze -- ukupno i dalje N Excel-a. Radi samo nad kopijom u stanju
    snimka: radnik dolazi posle mutacija, a kopiju vraca svaki prolaz -- osim
    onog koji javi REVERT-FAIL. Baze idu JEDNA PO JEDNA (jedna nit): pisu
    marker zelenog, a drugih pisaca markera tada nema.

    Otkaz ubija Excel baze po excel.pid (samo uz isto vreme nastanka), i to
    SAMO dok je radnik kod baze: posle vracanja isti izlaz koristi tudji posao."""

    def __init__(self, radnici, pre: str, ubij=None, otisak=None):
        self.suite = []
        self.rez = {}
        self._radnici = radnici
        self._pre = pre
        self._otisak = otisak or _otisak
        self._brava = threading.Lock()
        # Gotova baza (ili kraj posla) BUDI onoga ko ceka -- bez toga sacekaj
        # vidi napredak tek posle celog koraka cekanja.
        self._promena = threading.Condition(self._brava)
        self._otkazi = threading.Event()
        self._aktivan = False
        self._ctx = None
        self._ubij = ubij or (lambda c: _ubij_zaostali(c.get("izlaz")))
        self._ex = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        self._f = None
        _AKTIVNI.append(self._ex)
        _UNAPRED.append(self)

    def trazi(self, suite) -> list:
        """Stavi suite u red (svaku jednom) i vrati NOVE. Prva pokrece posao."""
        with self._brava:
            if self._otkazi.is_set():
                return []
            nove = [s for s in suite if s not in self.suite]
            self.suite.extend(nove)
            if nove and not self._aktivan:
                self._aktivan = True
                self._f = self._ex.submit(self._radi)
            return nove

    def _radi(self) -> None:
        zavrseno = False
        try:
            while True:
                with self._brava:
                    ostale = [s for s in self.suite if s not in self.rez]
                    if not ostale or self._otkazi.is_set() or _STOP.is_set():
                        # Odluka i gasenje su atomski: trazi() koji stigne posle
                        # ovoga vidi neaktivan posao i pokrece nov.
                        self._aktivan = False
                        self._promena.notify_all()
                        zavrseno = True
                        return
                c = self._radnici.uzmi(prioritet=True, odustani=lambda: (
                    self._otkazi.is_set() or _STOP.is_set()))
                if c is None:
                    continue
                try:
                    with self._brava:
                        self._ctx = c
                    if self._otisak(c["kopija"]) != self._pre:
                        rez = (False, "kopija radnika r%s nije u stanju snimka -- "
                                      "baza se nad njom ne meri" % c.get("id"))
                    else:
                        rez = _baza_zelena(ostale[0], src=c["kopija"],
                                           izlaz=c.get("izlaz"), prekini=self._otkazi)
                    with self._brava:
                        self.rez[ostale[0]] = rez
                        self._promena.notify_all()
                finally:
                    with self._brava:
                        self._ctx = None
                    self._radnici.vrati(c)
        finally:
            if not zavrseno:
                with self._brava:
                    self._aktivan = False
                    self._promena.notify_all()

    def gotovo(self) -> bool:
        """Nista ne radi i nista ne ceka (i kad nijedna baza nije trazena)."""
        with self._brava:
            return not self._aktivan

    def gotove(self) -> set:
        with self._brava:
            return set(self.rez)

    def rezultat(self, suite: str):
        """(ok, opis) baze koja je VEC gotova, ili None -- bez cekanja."""
        with self._brava:
            return self.rez.get(suite)

    def sacekaj(self, potrebne=None) -> dict:
        """Sacekaj baze koje faza 3 TREBA (u koracima -- Ctrl+C mora da stigne)
        i vrati gotove. Cim su potrebne gotove, ostatak se otkazuje: faza 3 ne
        ceka punu suite koja joj ne treba."""
        potrebne = set(self.suite if potrebne is None else potrebne) & set(self.suite)
        while True:
            with self._promena:
                if potrebne <= set(self.rez):
                    gotove = True
                elif not self._aktivan:
                    break
                else:
                    # U koracima: Ctrl+C na Windows-u ne stize u beskonacno cekanje.
                    self._promena.wait(0.5)
                    continue
            if gotove:
                self.otkazi()
                break
        f = self._f
        if f is not None and not self._otkazi.is_set():
            try:
                f.result(timeout=30)        # nit je na izlazu; izuzetak ide gore
            except concurrent.futures.TimeoutError:
                pass
        self._zatvori()
        with self._brava:
            return dict(self.rez)

    def otkazi(self) -> None:
        """Baze nisu potrebne: ne pokreci sledecu, a tekucu prekini."""
        self._otkazi.set()
        while True:
            with self._brava:
                if not self._aktivan:
                    break
                # Pod bravom: nit baze ne moze da vrati radnika dok se ubija.
                if self._ctx is not None:
                    self._ubij(self._ctx)
                f = self._f
            try:
                f.result(timeout=0.5)
            except concurrent.futures.TimeoutError:
                continue
            except Exception:               # noqa: BLE001 -- otkazana baza nije nalaz
                break
            break                           # posao je zavrsen -- nema sta da se ceka
        self._zatvori()

    def _zatvori(self) -> None:
        self._ex.shutdown(wait=True)
        if self._ex in _AKTIVNI:
            _AKTIVNI.remove(self._ex)
        if self in _UNAPRED:
            _UNAPRED.remove(self)


def _rana_mutacija(k: int, radnika: int, baza_radi: bool) -> bool:
    """Da li mutacije faze 3 krecu ODMAH, uz bazu koja jos radi: samo kad svih
    k staje na radnike koje baza ne drzi. Inace bi faza na N-1 radnika bila
    sporija od cekanja baze pa svih N, i baza se ceka pre mutacija."""
    return k <= radnika - (1 if baza_radi else 0)


# Ishod posla faze 3 koji se NE pusta: baza njegove suite je vec poznata i
# crvena, pa mutacija ne bi rekla nista. Serijski put je nikad ne pusta.
_PRESKOCENA = {"stop": "", "ishodi": [], "koren": None}


def _posao_faze3(s: dict, baza, prolaz):
    """Posao faze 3 za stavku s. baza(suite) -> (ok, opis) ili None (jos nije
    poznata), BEZ cekanja; prolaz(c) pusta mutaciju nad radnikom c.

    Baza se pita kad posao KRENE, ne kad se zakaze: uz radnike moze da postane
    poznata dok posao ceka slobodnog radnika."""
    def fn(c):
        b = baza(s["suite"])
        if b is not None and not b[0]:
            return dict(_PRESKOCENA)
        return prolaz(c)
    return fn


def _baza_poznata(cele_baze: dict, unapred, suite: str):
    """(ok, opis) baze koja je VEC poznata -- izmerena, ili gotova unapred --
    ili None, BEZ cekanja. unapred se cita JEDNOM: glavna nit ga posle prenosa
    postavlja na None, a njegov rez ostaje pun."""
    u = unapred
    b = cele_baze.get(suite)
    if b is None and u is not None:
        b = u.rezultat(suite)
    return b


def _baze_faze3_ovde(potrebne: set, unapred, cele_baze: dict) -> list:
    """Baze faze 3 koje se mere OVDE, serijski: potrebne, a niti vec izmerene
    niti u bazama unapred. Baza unapred se ne meri drugi put -- dve pune suite
    uporedo bi obe pisale marker zelenog."""
    bez = set(unapred.suite) if unapred is not None else set()
    return sorted(set(potrebne) - bez - set(cele_baze))


def _uzmi_baze_unapred(unapred, potrebne: set, cele_baze: dict, nacin: str) -> bool:
    """Prenesi baze unapred u cele_baze, ako smeju (v. _faza3_ocena). True kad
    su prenete -- unapred je tada zavrsen (sacekaj otkazuje nepotrebne).

    Bez cekanja ("proba", "stop") baze se uzimaju cim su POTREBNE gotove, i
    kad nepotrebna jos traje; gotov ceo posao se uzima uvek, da izuzetak iz
    baze ide gore kroz sacekaj."""
    spremne = potrebne & set(unapred.suite) <= unapred.gotove() or unapred.gotovo()
    if nacin != "cekaj" and not spremne:
        return False
    for suite, ok_opis in sorted(unapred.sacekaj(potrebne).items()):
        if suite in potrebne:
            cele_baze[suite] = ok_opis
            print("BAZNO (unapred): %s" % ok_opis[1], flush=True)
    return True


def _faza3_ocena(sve: list, tok, baze, upisi, lose: list, koren, baza) -> str:
    """Ocena mutacija cele suite (faza 3) -- ISTA za serijski put i za radnike.

    sve    stavke redom kataloga; tok daje njihove ishode ISTIM redom
           (None = prekid).
    baze   baze(nacin) -> {suite: (ok, opis)} ili None dok potrebne nisu
           poznate. nacin: "proba" (bez cekanja), "cekaj" (ceka potrebne),
           "stop" (samo vec gotove -- STOP ne ceka punu suite).
    upisi  zajednicko knjizenje ishoda; lose spisak problema; koren(ishod).
    baza   baza(suite) -> (ok, opis) baze koja je VEC poznata, ili None.

    Pravila, isto za oba puta:
    1. Stavke CRVENE baze javljaju se cim je baza poznata, SVE ODJEDNOM i
       redom kataloga -- pre prvog ishoda (serijski put ih je uvek tako
       javljao, pa izlaz N=1 i N>1 ostaje isti).
    2. Mutacija se priznaje SAMO uz zelenu bazu svoje suite; baza koje nema
       nije zelena.
    3. STOP ne ceka bazu koja jos traje: stavka je tada NIJE IZMERENA (dokaz je
       ionako NIJE DOKAZANO). Baza koja je VEC gotova se uzima i uz STOP -- i
       sopstvena baza stavke kad druga potrebna jos traje.

    Vraca razlog zaustavljanja ili ""."""
    def javi_crvene(b, koje=None):
        for s2 in (sve if koje is None else koje):
            ok2, opis2 = b.get(s2["suite"], (False, "baza cele suite nije izmerena"))
            if not ok2:
                lose.append((s2["ime"], "NIJE IZMERENA: ni izbor ni cela suite "
                                        "nisu zeleni -- %s" % opis2))
                print("%-46s NIJE IZMERENA: baza cele suite nije zelena"
                      % s2["ime"], flush=True)

    poznate = baze("proba")
    if poznate is not None:
        javi_crvene(poznate)
    for s, ishod in zip(sve, tok):
        if ishod is None:
            return "PREKID"
        koren(ishod)
        if poznate is None:
            poznate = baze("stop" if ishod["stop"] else "cekaj")
            if poznate is not None:
                javi_crvene(poznate)
        b = poznate.get(s["suite"], (False, "")) if poznate is not None else (
            baza(s["suite"]) if ishod["stop"] else None)
        if b is None:
            lose.append((s["ime"], "NIJE IZMERENA: ni izbor ni cela suite nisu "
                                   "zeleni -- baza cele suite nije izmerena "
                                   "(dokaz je stao pre nje)"))
            print("%-46s NIJE IZMERENA: dokaz je stao pre baze cele suite"
                  % s["ime"], flush=True)
        elif b[0]:
            upisi(ishod["ishodi"])
        elif poznate is None:
            # Sopstvena baza je crvena, a ostale nisu javljene: javi ovu sada.
            javi_crvene({s["suite"]: b}, [s])
        if ishod["stop"]:
            return ishod["stop"]
    return ""


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


# --- G1: prijem izbora -------------------------------------------------------
#
# Svaka stavka uzorka ide izborom (faza 1, kao inace), pa se verdikt izbora
# ZAPAMTI i stavka ipak meri celom suite-om (faza 3, nepromenjena). Na kraju se
# dva verdikta porede po stavci. Sve ispod je cisto -- main samo zove.

def _prijem_ishod(ishod):
    """Ishod faze 1 za prijem: verdikt izbora u "prijem", svaka stavka u
    "cela_suite" (tacno jednom), "ishodi" prazni -- konacni verdikt daje faza 3.

    prijem = [(stavka, priznato, stanje, vrsta)]; vrsta je 'izbor' (izbor je
    priznao), 'konacno' (izbor je presudio NE bez povratka: APPLY-FAIL,
    TIMEOUT, EXCEL UBIJEN) ili 'povratak' (izbor nije presudio -- stanje je
    razlog povratka). Ishod bez izbora ili sa stop-om vraca se NETAKNUT.

    Verdikt se hvata OVDE, pre knjizenja: kad stavka jednom stigne u
    cela_suite, njen konacni verdikt dolazi iz faze 3, pa bi poredjenje posle
    toga bilo cela suite prema samoj sebi (placebo)."""
    if not isinstance(ishod, dict) or not ishod.get("izbor") or ishod.get("stop"):
        return ishod
    prijem, cela = [], []
    for s, priznato, stanje, _crveno in ishod.get("ishodi", ()):
        prijem.append((s, priznato, stanje, "izbor" if priznato else "konacno"))
        cela.append((s, "prijem: izbor %s" % stanje))
    for s, razlog in ishod.get("cela_suite", ()):
        prijem.append((s, False, razlog, "povratak"))
        cela.append((s, "prijem: izbor %s" % razlog))
    novi = dict(ishod)
    novi.update(ishodi=[], cela_suite=cela, prijem=prijem)
    return novi


def _g1_hes(seme: str, x: str) -> str:
    return hashlib.sha256(("%s|%s" % (seme, x)).encode("utf-8")).hexdigest()


def _g1_otisak(imena) -> str:
    """Otisak uzorka: sha256 sortiranih imena, 16 hex."""
    return hashlib.sha256("\n".join(sorted(imena)).encode("utf-8")).hexdigest()[:16]


def _g1_uzorak(stavke: list, n: int, izbor_suite, s4a, posle, poznati_dokaz,
               poznati_staticki, seme: str = G1_SEME) -> tuple:
    """(sloj_po_imenu, problem) -- deterministicki uzorak prijema.

    Okvir su suite izbora (banka nema izbor, pa bi oba merenja bila ista).
    CILJANI sloj su sve stavke testova iz s4a i posle, i kljucevi poznati_dokaz
    (negativne kontrole). Ime sa spiska bez ijedne stavke je problem, nikad
    tiho preskakanje. SLUCAJNI sloj: okvir bez ciljanih i bez statickih
    POZNATI (njihov kljuc se ne poklapa ni u jednom rezimu), kvota po suiti
    srazmerna okviru; po suiti se fajlovi redjaju po sha256(seme|fajl), stavke
    u fajlu po sha256(seme|ime), bira se round-robin po fajlu, jedna po testu,
    i uzima prvih k. Sekvenca ne zavisi od k, pa je uzorak za veci n NADSKUP
    uzorka za manji. Izbor je clanstvo po imenu, nikad prefiks."""
    okvir = [s for s in stavke if s["suite"] in izbor_suite]
    ciljani_testovi = set(s4a) | set(posle)
    imena_okvira = {s["ime"] for s in okvir}
    testovi_okvira = {s["test"] for s in okvir}
    for test in sorted(ciljani_testovi - testovi_okvira):
        return {}, ("G1: ciljani test %s nema nijednu sabotazu u suite-ama "
                    "izbora -- ispravi G1_S4A_TESTOVI / G1_POSLE_IZBORA_TESTOVI" % test)
    for ime in sorted(set(poznati_dokaz) - imena_okvira):
        return {}, ("G1: POZNATI_NALAZI_DOKAZ['%s'] nema sabotazu u suite-ama "
                    "izbora" % ime)
    sloj = {}
    for s in okvir:
        if s["test"] in ciljani_testovi or s["ime"] in poznati_dokaz:
            sloj[s["ime"]] = "ciljani"
    van = set(poznati_staticki) - set(poznati_dokaz)
    bazen = [s for s in okvir if s["ime"] not in sloj and s["ime"] not in van]
    suite_redom = sorted({s["suite"] for s in okvir})
    dodeljeno = 0
    for i, suite in enumerate(suite_redom):
        u_okviru = sum(1 for s in okvir if s["suite"] == suite)
        k = (n - dodeljeno if i == len(suite_redom) - 1
             else int(round(n * u_okviru / float(len(okvir)))))
        dodeljeno += k
        po_fajlu = {}
        for s in bazen:
            if s["suite"] == suite:
                po_fajlu.setdefault(s["fajl"], []).append(s)
        redovi = [sorted(po_fajlu[f], key=lambda s: (_g1_hes(seme, s["ime"]), s["ime"]))
                  for f in sorted(po_fajlu, key=lambda f: (_g1_hes(seme, f), f))]
        poz, vidjeni, niz = [0] * len(redovi), set(), []
        while len(niz) < k:
            pomak = False
            for j, red in enumerate(redovi):
                while poz[j] < len(red):
                    s = red[poz[j]]
                    poz[j] += 1
                    if s["test"] not in vidjeni:
                        vidjeni.add(s["test"])
                        niz.append(s)
                        pomak = True
                        break
            if not pomak:
                break
        if len(niz) < k:
            return {}, ("G1: %s ima %d razlicitih testova van ciljanih, a kvota "
                        "je %d" % (suite, len(niz), k))
        for s in niz[:k]:
            sloj[s["ime"]] = "slucajni"
    return sloj, ""


def _g1_po_suiti(suite_spisak) -> str:
    """'RAT 34, BFP 9' -- redom RAT, BFP, banka, pa ostale."""
    broj = {}
    for suite in suite_spisak:
        broj[suite] = broj.get(suite, 0) + 1
    redom = [s for s in (SUITE_ALL, SUITE_BFP, SUITE_BANKA) if s in broj]
    redom += sorted(set(broj) - set(redom))
    return ", ".join("%s %d" % (_G1_KRATKO.get(s, s), broj[s]) for s in redom)


def _g1_uzorak_red(stavke: list, sloj: dict, presudjeni=None) -> str:
    """Red G1 UZORAK: otisak i velicina slojeva (i presudjeni, na kraju)."""
    delovi = []
    for ime_sloja, rec in (("ciljani", "ciljanih"), ("slucajni", "slucajnih"),
                           ("filter", "filter")):
        clanovi = [s["suite"] for s in stavke if sloj.get(s["ime"]) == ime_sloja]
        if clanovi:
            delovi.append("%s %d (%s)" % (rec, len(clanovi), _g1_po_suiti(clanovi)))
    red = "G1 UZORAK: otisak %s -- %s" % (
        _g1_otisak(s["ime"] for s in stavke if s["ime"] in sloj), ", ".join(delovi))
    if presudjeni is not None:
        red += "; presudjeno izborom %d (%s)" % (
            len(presudjeni), _g1_po_suiti(r["suite"] for r in presudjeni) or "-")
    return red


def _g1_infra(stanje: str) -> str:
    return next((p for p in _G1_INFRA if stanje.startswith(p)), "")


def _g1_nemereno(stanje: str) -> bool:
    """Strana koja tvrdnju NIJE posmatrala: sidro se nije primenilo, ili prolaz
    nije stigao do kraja. Jednakost dve takve strane nije poredjenje."""
    return stanje.startswith("APPLY-FAIL") or bool(_g1_infra(stanje))


# Razlog crvene baze izbora je "<suite>: <opis>" (_baza_zelena). Ovi opisi kazu
# da baza NIJE ODRADJENA, ne da test sam ne prolazi -- pa nisu S-4a regresija.
_G1_BAZA_INFRA = ("timeout", "EXCEL UBIJEN", "nema oznacenog reda TESTS")


def _g1_klasa(izb, cela) -> tuple:
    """(klasa, podtip) jedne stavke. izb = (priznato, stanje, vrsta) iz
    _prijem_ishod ili None; cela = (priznato, stanje) iz faze 3 ili None.
    Poredi se PRIZNANJE (bool), ne tekst stanja: "OK (uz jos 2)" je OK."""
    if izb is None or cela is None:
        return "NEUPOREDIVO", ""
    i_ok, i_st, vrsta = izb
    c_ok, c_st = cela
    if vrsta == "povratak":
        if not i_st.startswith("u izboru:"):
            opis = i_st.split(": ", 1)[-1]
            if opis.startswith(_G1_BAZA_INFRA):
                return "SKUPA", "baza izbora infra"
            return "SKUPA", "baza izbora crvena"
        return "SKUPA", "cela OK" if c_ok else "cela NE"
    if vrsta == "izbor":
        if c_ok:
            return "ISTO", ""
        if _g1_infra(c_st):
            return "NEJASNO", ""
        if c_st.startswith("NE OBARA NISTA"):
            return "OPASNA", "MASKA"
        return "OPASNA", "DRUGI PAD"
    if vrsta != "konacno":
        return "NEUPOREDIVO", ""
    # 'konacno' znaci da izbor tvrdnju NIJE posmatrao (APPLY-FAIL, TIMEOUT,
    # EXCEL UBIJEN), pa ISTO ovde ne postoji. Ni cela suite nije stigla do
    # tvrdnje: par je NEUPOREDIV i ne ulazi u imenilac -- jednakost bez merenja
    # je placebo iz istog razloga kao SKUPA (review G1, 17.11.2026).
    if not c_ok and _g1_nemereno(c_st):
        return "NEUPOREDIVO", "nijedna strana ne meri"
    if _g1_infra(i_st) or not c_ok:
        return "NEJASNO", ""
    return "STROZA", ""


def _g1_poredi(stavke: list, prijem: list, verdikt: dict, sloj: dict) -> list:
    """Redovi poredjenja, redom kataloga, za svaku stavku uzorka."""
    izb = {s["ime"]: (p, st, v) for s, p, st, v in prijem}
    redovi = []
    for s in stavke:
        if s["ime"] not in sloj:
            continue
        i, c = izb.get(s["ime"]), verdikt.get(s["ime"])
        klasa, podtip = _g1_klasa(i, c)
        redovi.append({"ime": s["ime"], "suite": s["suite"], "test": s["test"],
                       "sloj": sloj[s["ime"]], "klasa": klasa, "podtip": podtip,
                       "izbor": i, "cela": c})
    return redovi


def _g1_presudjeni(redovi: list) -> list:
    """Imenilac: SLUCAJNE stavke koje je izbor presudio sam, uz izmerenu celu
    suite. SKUPA (placebo), NEJASNO i NEUPOREDIVO ne ulaze."""
    return [r for r in redovi if r["sloj"] == "slucajni"
            and r["izbor"] is not None and r["izbor"][2] in ("izbor", "konacno")
            and r["cela"] is not None
            and r["klasa"] in ("ISTO", "OPASNA", "STROZA")]


def _g1_presuda(redovi: list, stop: str, celovit: bool, n_min: int,
                suite_min: int) -> tuple:
    """(rc, red G1 PRIJEM). 1 PAO: bar jedna OPASNA ili STROZA, u bilo kom
    sloju. 2 NEJASNO: STOP, snimak se menjao, NEJASNO stavka, ciljana
    NEUPOREDIVO, premalo presudjenih (ukupno ili u RAT/BFP), filter. 0 PROSAO."""
    presudjeni = _g1_presudjeni(redovi)
    m = len(presudjeni)
    razlike = [r for r in redovi if r["klasa"] in ("OPASNA", "STROZA")]
    razlika_slucajnih = sum(1 for r in presudjeni if r["klasa"] != "ISTO")
    # Udeo povrataka PO SLOJU: samo slucajni procenjuje trosak podrazumevanog
    # izbora (D-2, "[?]"); ciljani je namerno pristrasan (POZNATI su negativne
    # kontrole, ocekuje se SKUPA).
    skupa = ", ".join(
        "%s %d/%d" % (sl, sum(1 for r in redovi if r["sloj"] == sl
                              and r["klasa"] == "SKUPA"),
                      sum(1 for r in redovi if r["sloj"] == sl))
        for sl in ("slucajni", "ciljani", "filter")
        if any(r["sloj"] == sl for r in redovi))
    nejasno = []
    if stop:
        nejasno.append("dokaz je stao (%s)" % stop.split(":")[0])
    elif not celovit:
        nejasno.append("snimak pre/posle se razlikuje")
    imena = [r["ime"] for r in redovi if r["klasa"] == "NEJASNO"]
    if imena:
        nejasno.append("NEJASNO: %s" % ", ".join(imena))
    imena = [r["ime"] for r in redovi
             if r["sloj"] == "ciljani" and r["klasa"] == "NEUPOREDIVO"]
    if imena:
        nejasno.append("ciljani NEUPOREDIVO: %s" % ", ".join(imena))
    if any(r["sloj"] == "filter" for r in redovi):
        nejasno.append("poziv sa filterom nije uzorak")
    else:
        if m < n_min:
            nejasno.append("presudjeno izborom %d < %d (npr. --prijem 75)" % (m, n_min))
        for suite in (SUITE_ALL, SUITE_BFP):
            k = sum(1 for r in presudjeni if r["suite"] == suite)
            if k < suite_min:
                nejasno.append("presudjeno u %s %d < %d"
                               % (_G1_KRATKO[suite], k, suite_min))
    if razlike:
        rc, ishod = 1, "PAO"
        razlog = "; ".join(
            "%s: %s" % (k, ", ".join(r["ime"] for r in razlike if r["klasa"] == k))
            for k in ("OPASNA", "STROZA") if any(r["klasa"] == k for r in razlike))
    elif nejasno:
        rc, ishod, razlog = 2, "NEJASNO", " / ".join(nejasno)
    else:
        rc, ishod, razlog = 0, "PROSAO", "nijedna razlika"
    if m and not razlika_slucajnih:
        granica = "van ciljanih najvise %s%% uz 95%%" % (
            ("%.1f" % (100.0 * (1.0 - 0.05 ** (1.0 / m)))).replace(".", ","))
    else:
        granica = "granica ne vazi"
    return rc, ("G1 PRIJEM: %s -- %s; %d razlika nad %d presudjenih slucajnih; "
                "%s; SKUPA %s" % (ishod, razlog, razlika_slucajnih, m, granica,
                                  skupa))


def _g1_ispis(redovi: list, stavke: list, sloj: dict) -> list:
    """Redovi izvestaja pre G1 PRIJEM: svaka stavka, KLASE, S-4a, UZORAK."""
    out = []
    for r in redovi:
        i, c = r["izbor"], r["cela"]
        out.append("G1 %s%s %s %s %s | izbor: %s%s | cela: %s" % (
            r["klasa"], "/" + r["podtip"] if r["podtip"] else "", r["sloj"],
            _G1_KRATKO.get(r["suite"], r["suite"]), r["ime"],
            i[1] if i else "nije izmereno",
            " -> cela suite" if i and i[2] == "povratak" else "",
            c[1] if c else "nije izmerena"))
        if r["klasa"] not in ("ISTO", "SKUPA"):
            out.append("  ponovi: python tools/dokaz.py --prijem %s --radnika 1"
                       % r["ime"])

    def broj(k, podtip=None):
        return sum(1 for r in redovi if r["klasa"] == k
                   and (podtip is None or r["podtip"] in podtip))
    out.append("G1 KLASE: ISTO %d, SKUPA %d (izbor ne obara %d, baza izbora "
               "crvena %d, baza izbora infra %d), OPASNA %d, STROZA %d, "
               "NEJASNO %d, NEUPOREDIVO %d" % (
                   broj("ISTO"), broj("SKUPA"),
                   broj("SKUPA", ("cela OK", "cela NE")),
                   broj("SKUPA", ("baza izbora crvena",)),
                   broj("SKUPA", ("baza izbora infra",)), broj("OPASNA"),
                   broj("STROZA"), broj("NEJASNO"), broj("NEUPOREDIVO")))
    for test in sorted({r["test"] for r in redovi if r["test"] in G1_S4A_TESTOVI
                        and r["podtip"] == "baza izbora crvena"}):
        out.append("G1 S-4a REGRESIJA: %s" % test)
    out.append(_g1_uzorak_red(stavke, sloj, _g1_presudjeni(redovi)))
    return out


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
    ap.add_argument("--prijem", nargs="?", const=str(G1_N), default=None,
                    metavar="N",
                    help="G1 prijem izbora: svaka stavka uzorka (ciljani + N "
                         "slucajnih, podrazumevano %d) ide izborom PA celom "
                         "suite-om, i verdikti se porede; uz filter meri samo "
                         "filtrirane stavke (presuda je tada NEJASNO)" % G1_N)
    ap.add_argument("--self-test", action="store_true",
                    help="dokazi da pravila grupisanja, ocena i orkestracija prolaza grizu")
    args = ap.parse_args(argv)
    # `--prijem ime`: argparse uzme ime kao N (nargs='?'), pa se vraca filteru.
    # Bez opcije je None (ugaseno). Izricito 0 se ODBIJA (dole): da je 0 isto
    # sto i izostavljena opcija, `--prijem 0` bi tiho pustio pun obican dokaz
    # celog kataloga, bez ijednog G1 reda.
    prijem_nula = False
    if args.prijem is None:
        args.prijem = 0
    elif str(args.prijem).isdigit():
        args.prijem = int(args.prijem)
        prijem_nula = args.prijem == 0
    else:
        args.filter.insert(0, args.prijem)
        args.prijem = G1_N

    if args.self_test:
        return _self_test()
    if args.grupe < 1:
        print("--grupe mora biti >= 1", file=sys.stderr)
        return 2
    if args.radnika < 1:
        print("--radnika mora biti >= 1", file=sys.stderr)
        return 2
    if prijem_nula:
        print("--prijem 0: prijem bez slucajnih ne postoji -- izostavi opciju "
              "ili zadaj N >= %d" % G1_MIN, file=sys.stderr)
        return 2
    if args.prijem:
        if args.grupe > 1 or args.izolacija:
            print("--prijem meri svaku stavku POJEDINACNO, izborom pa celom "
                  "suite-om -- ne ide uz --grupe > 1 ni uz --izolacija",
                  file=sys.stderr)
            return 2
        if args.prijem < G1_MIN and not args.filter:
            print("--prijem %d: manje od %d slucajnih ne moze da da PROSAO "
                  "(G1_MIN)" % (args.prijem, G1_MIN), file=sys.stderr)
            return 2
        args.izbor = True
    _STOP.clear()
    _START.update(poslednji=0.0, ponovljeno=0,
                  razmak=args.razmak if args.radnika > 1 else 0.0)

    sab = _modul_sabotaza()
    poznati_spisak = getattr(sab, "POZNATI_NALAZI_DOKAZ", {})
    stavke = _izaberi(sab, args.filter)

    if not stavke:
        print("filter ne pogadja nijednu sabotazu", file=sys.stderr)
        return 2

    sloj = {}
    if args.prijem and args.filter:
        sloj = {s["ime"]: "filter" for s in stavke}
    elif args.prijem:
        sloj, problem = _g1_uzorak(
            stavke, args.prijem, _suite_izbor(), G1_S4A_TESTOVI,
            G1_POSLE_IZBORA_TESTOVI, poznati_spisak,
            set(getattr(sab, "POZNATI_NALAZI", {})) - set(poznati_spisak), G1_SEME)
        if problem:
            print("STOP: " + problem, file=sys.stderr)
            return 2
        stavke = [s for s in stavke if s["ime"] in sloj]

    if args.plan:
        if not args.prijem:
            return _plan(stavke, args.grupe)
        for s in stavke:
            print("G1 %-8s %-5s %s" % (sloj[s["ime"]],
                                       _G1_KRATKO.get(s["suite"], s["suite"]), s["ime"]))
        rc = _plan(stavke, args.grupe)
        print(_g1_uzorak_red(stavke, sloj))
        return rc

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
    if args.prijem:
        stavke = [s for s in stavke if s["ime"] in sloj]
        if len(stavke) != len(sloj):
            print("STOP: G1: nad snimkom je %d stavki uzorka, a plan ih ima %d "
                  "-- katalog se menjao" % (len(stavke), len(sloj)), file=sys.stderr)
            shutil.rmtree(ctx["radni"], ignore_errors=True)
            return 2
        print(_g1_uzorak_red(stavke, sloj), flush=True)

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

    # Radnici su ZAJEDNICKI za sve faze i za baze unapred (v. _Radnici).
    radnici = _Radnici(bazen)

    # Baze cele suite UNAPRED (v. _BazeUnapred): samo uz radnike i izbor, i to
    # NA ZAHTEV -- krecu uz prvu stavku koja vodi u fazu 3 (najavi).
    unapred = None
    moguce = set()
    if args.radnika > 1 and izbor:
        moguce = ({s["suite"] for s in stavke} & izbor) - set(cele_baze)
        if moguce:
            unapred = _BazeUnapred(radnici, pre)

    def najavi(ishod):
        """Stavka koja SIGURNO ide u fazu 3 (cela suite) trazi bazu svoje suite
        ODMAH, iz niti radnika -- ne ceka redosled ispisa. Ponavljanje sama se
        ne racuna: ta stavka cesto prodje u fazi 2, a njena baza bi stala u red
        ISPRED baze koja fazi 3 stvarno treba (baze idu jedna po jedna).

        G1 (--prijem): verdikt izbora se hvata OVDE, pre svakog knjizenja, a
        svaka stavka ide u fazu 3 -- pa baze unapred krecu uz prvu stavku."""
        if args.prijem:
            ishod = _prijem_ishod(ishod)
        if unapred is not None and isinstance(ishod, dict):
            kandidati = [s for s, _r in ishod.get("cela_suite", [])]
            nove = unapred.trazi(sorted({s["suite"] for s in kandidati} & moguce))
            if nove:
                print("BAZE UNAPRED: %s -- krece uz %s (prva stavka koja vodi u "
                      "fazu 3); otkazuje se ako ne zatreba"
                      % (", ".join(nove), kandidati[0]["ime"]), flush=True)
        return ishod

    crvenih, lose, poznati = 0, [], []
    grupno_priznatih, izborom_priznatih, stop = 0, 0, ""
    verdikt, prijem_izbor = {}, []    # G1: sirov verdikt po imenu; verdikti izbora

    def upisi(ishodi, uvlaka=""):
        """Zajednicki ispis i knjizenje za grupni i pojedinacni prolaz."""
        nonlocal crvenih
        for s, priznato, stanje, crveno in ishodi:
            verdikt[s["ime"]] = (priznato, stanje)
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

    def posao(clanovi, grupno, izb=frozenset(), baz=None, najava=False):
        if najava:
            return lambda c: najavi(_prolaz(clanovi, pre, grupno, izb, baz, c))
        return lambda c: _prolaz(clanovi, pre, grupno, izb, baz, c)

    with contextlib.closing(_izvrsi([posao(g, len(g) > 1, izbor, baze_izbora, True)
                                     for g in grupe], radnici)) as tok:
        for i, (g, ishod) in enumerate(zip(grupe, tok), 1):
            if ishod is None:
                stop = stop or "PREKID"
                break
            if len(g) > 1:
                print("GRUPA %d/%d  %s  clanova %d"
                      % (i, len(grupe), g[0]["suite"], len(g)), flush=True)
            koren_tokom.add(ishod.get("koren") or koren_pre)
            cela_suite.extend(ishod["cela_suite"])
            prijem_izbor.extend(ishod.get("prijem", ()))
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
        with contextlib.closing(_izvrsi([posao([s], False, izbor, baze_izbora, True)
                                         for s in ponavljaju], radnici)) as tok:
            for s, ishod in zip(ponavljaju, tok):
                if ishod is None:
                    stop = stop or "PREKID"
                    break
                koren_tokom.add(ishod.get("koren") or koren_pre)
                cela_suite.extend(ishod["cela_suite"])
                prijem_izbor.extend(ishod.get("prijem", ()))
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
        print(("MERI SE CELOM SUITE-OM: %d (prijem: svaka stavka ide i celom "
               "suite-om; razlog uz ime je verdikt izbora)" if args.prijem else
               "MERI SE CELOM SUITE-OM: %d (ciljani test sam ne prolazi, ili izbor "
               "nije oborio tvrdnju -- v. --izolacija)") % len(cela_suite), flush=True)
        for s, razlog in cela_suite:
            print("  %-46s <- %s" % (s["ime"], razlog), flush=True)
        potrebne_3 = {s["suite"] for s, _r in cela_suite}
        # Baze koje niko ne meri unapred mere se ovde, serijski (pisu marker).
        for suite in _baze_faze3_ovde(potrebne_3, unapred, cele_baze):
            # Radnik se POZAJMLJUJE: baza unapred moze da drzi bilo kog.
            c = radnici.uzmi(prioritet=True)
            try:
                ok_opis = _baza_zelena(suite, src=c["kopija"], izlaz=c["izlaz"])
            finally:
                radnici.vrati(c)
            cele_baze[suite] = ok_opis
            print("BAZNO: %s" % cele_baze[suite][1], flush=True)

        def baze(nacin):
            """Baze za ocenu faze 3 (v. _faza3_ocena): bez baza unapred su vec
            izmerene; uz njih se uzimaju tek kad su potrebne gotove, osim uz
            "cekaj"."""
            nonlocal unapred
            if unapred is not None:
                if not _uzmi_baze_unapred(unapred, potrebne_3, cele_baze, nacin):
                    return None
                unapred = None
            return cele_baze

        def baza_poznata(suite):
            return _baza_poznata(cele_baze, unapred, suite)

        sve = [s for s, _r in cela_suite]
        # RANA MUTACIJA (R-4): uz baze unapred mutacije cele suite krecu ODMAH,
        # dok baza jos traje, a baza se ceka tek pri oceni -- kraj dokaza je
        # max(baza, mutacija), ne zbir. Samo kad sve stavke staju na radnike
        # koje faze imaju BEZ rezervisanog; inace bi N-1 radnika bilo sporije
        # od cekanja baze pa svih N, i baza se ceka pre mutacija, kao ranije.
        if not _rana_mutacija(len(sve), len(radnici),
                              unapred is not None and not unapred.gotovo()):
            baze("cekaj")
        with contextlib.closing(_izvrsi(
                [_posao_faze3(s, baza_poznata,
                              posao([s], False)) for s in sve],
                radnici)) as tok:
            stop = _faza3_ocena(
                sve, tok, baze, upisi, lose,
                lambda ishod: koren_tokom.add(ishod.get("koren") or koren_pre),
                baza_poznata)
        if stop and stop != "PREKID":
            _STOP.set()

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

    if unapred is not None and not unapred.suite:
        unapred.otkazi()                # nije ni trazena: nema sta da se javi
        unapred = None
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
    if izbor and args.prijem:
        print("izbor: baza i mutacija nad istim spiskom testova; prijem meri "
              "celom suite-om svih %d" % len(cela_suite))
    elif izbor:
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
    if args.prijem:
        # Red ispod je uzgredni verdikt OBICNOG dokaza nad uzorkom (cela suite);
        # bez ove napomene "NIJE DOKAZANO" uz "G1 PRIJEM: PROSAO" izgleda kao
        # protivrecnost (operater, 18.11.2026).
        print("(red ispod je verdikt obicnog dokaza nad uzorkom; za G1 vazi "
              "samo red G1 PRIJEM)")
    print("=== %s ===" % _verdikt(ok, grupno_priznatih, izborom_priznatih,
                                 pre if koren_menjan else ""))
    rc_g1 = None
    if args.prijem:
        # G1: verdikt izbora (faza 1) prema celoj suiti (faza 3), po stavci.
        celovit = not stop and pre == posle
        redovi = _g1_poredi(stavke, prijem_izbor, verdikt, sloj)
        rc_g1, red_g1 = _g1_presuda(redovi, stop, celovit, G1_MIN, G1_MIN_SUITE)
        for linija in _g1_ispis(redovi, stavke, sloj):
            print(linija)
        print(red_g1, flush=True)
    sacuvana = _zavrsi_istoriju(bazen, args.radnika)
    if stop or not sacuvana:
        print("radni folder ostaje: %s" % ctx["radni"], file=sys.stderr)
    else:
        shutil.rmtree(ctx["radni"], ignore_errors=True)
    return rc_g1 if args.prijem else (0 if ok else 1)


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
                  ctx=None, prekid=False, zatecen="P"):
    """Pokreni `_prolaz` bez Excela. Vraca (rez, pozivi).

    Pravila grupisanja i ocena su ciste funkcije i mere se direktno, ali
    ORKESTRACIJA nije: primeni sve (u kopiju) -> pusti suite TACNO JEDNOM ->
    obnovi kopiju iz snimka -> oceni, pa clan koji nije priznat vrati na
    pojedinacno merenje. Bez ovoga bi
    taj redosled prvi put bio izvrsen nad pravim Excelom, gde jedan prolaz traje
    minutima i gde greska u njemu izgleda kao nalaz nad VBA kodom.
    """
    pozivi = []
    stanje = {"mutirano": False, "obnovljena": False}

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
        stanje["obnovljena"] = True

    # Potpis ZAVISI OD PUTANJE: `potpis` vazi samo za kopiju, sve ostalo
    # (src-vba, snimak) je "P". Provera koja bi merila pogresno stablo -- npr.
    # src-vba umesto kopije -- tako bi videla "P" i propustila nevracenu kopiju.
    # I OD TRENUTKA: kopija PRE prolaza nosi `zatecen` (cista = "P"), a posle
    # obnove `potpis` -- nevracena kopija je ishod obnove, ne pocetno stanje.
    kopija = (ctx or _LAZNI_CTX)["kopija"]
    pravi = _pusti, _otisak, _obnovi
    globals()["_pusti"] = lazni_pusti
    globals()["_otisak"] = lambda put=None: (
        (potpis if stanje["obnovljena"] else zatecen) if put == kopija else "P")
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
    # Radnik iz zajednickog bazena moze da stigne sa TUDJOM, nevracenom
    # mutacijom (R-4): prolaz tada ne meri nista -- ni bazu izbora, ni mutaciju.
    rez, pozivi = prolaz(g2, OBA, zatecen="TUDJA", izbor=frozenset({"RunAllTests"}))
    tvrdi(rez["stop"].startswith("STRAZA") and "nije u stanju snimka" in rez["stop"]
          and not pozivi and not rez["ishodi"],
          "D-6: prolaz meri nad kopijom koja nije u stanju snimka (tudja mutacija "
          "iz zajednickog bazena)")

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
    for ime, igle in (("_prolaz(", ("baz, c)",)),
                      ("_baza_zelena(", ('src=ctx["kopija"]', 'src=c["kopija"]')),
                      ("_izolacija(", ('ctx["kopija"]',))):
        redovi = [r for r in izvor_main.splitlines() if ime in r]
        tvrdi(redovi and all(any(i in r for i in igle) for r in redovi),
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

    # --- baze unapred: NA ZAHTEV, na pozajmljenom radniku; otkaz ne ceka ---
    # Scenariji idu U ROKU: radnik koji se ne vrati, ili cekanje koje ne
    # odustaje, zaglavili bi self-test (i hook) -- mora da padne po imenu.
    def u_roku(fn, rok=3.0):
        rez = {}

        def telo():
            rez["v"] = fn()

        nit = threading.Thread(target=telo, daemon=True)
        nit.start()
        nit.join(rok)
        return (not nit.is_alive()), rez.get("v")

    redosled = []
    pravi_bz = _baza_zelena
    r5 = {"id": 5, "kopija": "/k5", "izlaz": "/i5"}
    cist = lambda _p: "P"          # noqa: E731 -- otisak kopije u stanju snimka

    def lazna(suite, testovi=None, src=None, izlaz=None, prekini=None):
        redosled.append((suite, src, izlaz))
        return True, "TESTS %s: 1 ukupno, 0 palo" % suite

    prvi = {}

    class Belezi(_Radnici):
        """Belezi prioritet zahteva i stanje baze u trenutku vracanja radnika."""
        def __init__(self, bazen):
            super().__init__(bazen, korak=0.01)
            self.prioriteti, self.ctx_pri_vracanju, self.bu = [], [], None

        def uzmi(self, prioritet=False, odustani=None, kljuc=None):
            self.prioriteti.append(prioritet)
            return super().uzmi(prioritet, odustani, kljuc)

        def vrati(self, c):
            if self.bu is not None:
                self.ctx_pri_vracanju.append(self.bu._ctx)
            super().vrati(c)

    def prvi_scenario():
        rad = prvi["rad"] = Belezi([r5])
        bu = prvi["bu"] = _BazeUnapred(rad, "P", ubij=lambda c: None, otisak=cist)
        rad.bu = bu
        time.sleep(0.02)
        prvi["nista"] = redosled == [] and bu.gotovo()
        prvi["nove"] = bu.trazi(["S1", "S2"])
        prvi["dupla"] = bu.trazi(["S1"])
        prvi["dobijeno"] = bu.sacekaj()
        prvi["vracen"] = rad.uzmi(odustani=lambda: True) is r5

    globals()["_baza_zelena"] = lazna
    try:
        gotov1, _v = u_roku(prvi_scenario)
    finally:
        if "bu" in prvi:
            prvi["bu"]._otkazi.set()    # pusta nit ako je zaglavljena
        globals()["_baza_zelena"] = pravi_bz
    nista_bez_zahteva = prvi.get("nista", False)
    nove, dupla = prvi.get("nove"), prvi.get("dupla")
    dobijeno, vracen = prvi.get("dobijeno") or {}, prvi.get("vracen") and gotov1
    bu = prvi.get("bu")
    tvrdi(nista_bez_zahteva,
          "baza unapred krece bez zahteva (rez bez faze 3 bi platio punu suite)")
    rad_p = prvi.get("rad")
    tvrdi(rad_p is not None and rad_p.prioriteti[:2] == [True, True],
          "baza unapred ne trazi radnika PRIORITETNO -- ceka iza poslova faza "
          "(%r)" % (rad_p.prioriteti if rad_p else None))
    tvrdi(rad_p is not None and len(rad_p.ctx_pri_vracanju) == 2
          and all(x is None for x in rad_p.ctx_pri_vracanju),
          "baza vraca radnika dok je jos 'njen' -- otkaz u tom trenutku ubija "
          "Excel faze koja ga uzme")
    tvrdi(redosled == [("S1", "/k5", "/i5"), ("S2", "/k5", "/i5")]
          and nove == ["S1", "S2"] and dupla == []
          and sorted(dobijeno) == ["S1", "S2"] and bu is not None
          and bu._ex not in _AKTIVNI and bool(vracen),
          "baze unapred se ne mere redom i jednom na pozajmljenom radniku (%r), "
          "radnik se ne vraca, ili izvrsilac ostaje prijavljen" % redosled)

    # Zahtev POSLE zavrsenog posla pokrece nov posao (nijedna suite ne ostaje).
    redosled[:] = []
    stanje = {}

    def zahtev_posle():
        bu = _BazeUnapred(_Radnici([r5], korak=0.01), "P", ubij=lambda c: None,
                          otisak=cist)
        stanje["bu"] = bu
        bu.trazi(["S1"])
        for _ in range(100):
            if bu.gotovo():
                break
            time.sleep(0.01)
        bu.trazi(["S2"])
        return bu.sacekaj()

    globals()["_baza_zelena"] = lazna
    try:
        gotov, dobijeno = u_roku(zahtev_posle)
    finally:
        if "bu" in stanje:
            stanje["bu"]._otkazi.set()  # pusta nit ako je zaglavljena
        globals()["_baza_zelena"] = pravi_bz
    tvrdi(gotov and sorted(dobijeno or {}) == ["S1", "S2"]
          and [x[0] for x in redosled] == ["S1", "S2"],
          "zahtev posle zavrsene baze se gubi, ili baza visi jer se radnik ne "
          "vraca (%r)" % redosled)

    # Kopija radnika van stanja snimka: baza se NE meri nad njom.
    redosled[:] = []
    mut = {}

    def van_snimka():
        bu = mut["bu"] = _BazeUnapred(_Radnici([r5], korak=0.01), "P",
                                      ubij=lambda c: None, otisak=lambda _p: "MUTIRANA")
        bu.trazi(["S1"])
        return bu.sacekaj()

    globals()["_baza_zelena"] = lazna
    try:
        gotov_m, dobijeno = u_roku(van_snimka)
        dobijeno = dobijeno or {}
    finally:
        if "bu" in mut:
            mut["bu"]._otkazi.set()
        globals()["_baza_zelena"] = pravi_bz
    tvrdi(gotov_m and redosled == [] and not dobijeno.get("S1", (True,))[0]
          and "nije u stanju snimka" in dobijeno.get("S1", (True, ""))[1],
          "baza se meri nad kopijom radnika koja nije u stanju snimka (%r)" % (dobijeno,))

    # Izuzetak iz baze unapred ide GORE kroz sacekaj, kao iz serijske baze
    # (N=1): progutan bi uz N>1 postao NIJE IZMERENA, a N=1 bi pukao.
    def pukne(suite, testovi=None, src=None, izlaz=None, prekini=None):
        raise RuntimeError("baza pukla")

    puk = {}

    def puknuta():
        bu = puk["bu"] = _BazeUnapred(_Radnici([r5], korak=0.01), "P",
                                      ubij=lambda c: None, otisak=cist)
        bu.trazi(["S1"])
        try:
            bu.sacekaj()
        except RuntimeError as e:
            return str(e)
        return None

    globals()["_baza_zelena"] = pukne
    try:
        gotov_p, pukla = u_roku(puknuta)
    finally:
        if "bu" in puk:
            puk["bu"]._otkazi.set()
        globals()["_baza_zelena"] = pravi_bz
    tvrdi(gotov_p and pukla == "baza pukla",
          "izuzetak iz baze unapred se guta (uz N>1 bi postao NIJE IZMERENA, a "
          "N=1 puca) (%r)" % (pukla,))

    # Otkaz: tekuca suite se prekida (ubija se SAMO dok je radnik kod baze),
    # sledeca ne krece, a radnik se vraca.
    redosled[:] = []
    ubijeno, pusti = [], threading.Event()

    def visi(suite, testovi=None, src=None, izlaz=None, prekini=None):
        tvrdi(prekini is not None, "baza unapred ne predaje svoj otkaz pozivu")
        redosled.append(suite)
        pusti.wait(2.0)
        return False, "ubijen"

    globals()["_baza_zelena"] = visi
    try:
        rad = _Radnici([r5], korak=0.01)
        vracen_tokom = []

        def ubij_visi(c):
            # Ubija se POD BRAVOM: baza ne sme da vrati radnika dok traje ubijanje.
            ubijeno.append(c["id"])
            pusti.set()
            time.sleep(0.05)
            vracen_tokom.append(r5 in rad._slobodni)

        bu = _BazeUnapred(rad, "P", ubij=ubij_visi, otisak=cist)
        bu.trazi(["S1", "S2"])
        time.sleep(0.02)
        u_registru = bu in _UNAPRED
        t0 = time.time()
        bu.otkazi()
        trajalo = time.time() - t0
        vracen = rad.uzmi(odustani=lambda: True) is r5
    finally:
        globals()["_baza_zelena"] = pravi_bz
    tvrdi(u_registru and bu not in _UNAPRED,
          "baza unapred nije u registru dok radi (prekid je ne bi otkazao), "
          "ili ostaje u njemu posle otkaza")
    tvrdi(vracen_tokom == [False],
          "otkaz ubija van brave: baza vraca radnika dok se ubija, pa ubijanje "
          "pogadja tudji posao na istom izlazu (%r)" % vracen_tokom)
    tvrdi(redosled == ["S1"] and ubijeno == [5] and trajalo < 1.5 and vracen,
          "otkaz baze unapred ne prekida tekucu suite, pusta sledecu, ili ne "
          "vraca radnika (%r, %.2fs)" % (redosled, trajalo))

    # Otkaz POSLE vracanja radnika ne ubija nista: isti izlaz je sada tudji.
    # Prozor je stvaran: baza je vratila radnika posle S1 i CEKA sledeceg za
    # S2, dok su svi radnici kod faza.
    class Zauzeti(_Radnici):
        def __init__(self, bazen):
            super().__init__(bazen, korak=0.01)
            self.zahteva = 0
            self.visi = False

        def uzmi(self, prioritet=False, odustani=None, kljuc=None):
            self.zahteva += 1
            if self.zahteva == 2:
                # U roku: cekanje koje ne odustaje na otkaz pada po imenu,
                # ne zaglavi self-test.
                rok = time.time() + 1.0
                while not (odustani is not None and odustani()):
                    if time.time() > rok:
                        self.visi = True
                        break
                    time.sleep(0.005)
                return None
            return super().uzmi(prioritet, odustani, kljuc)

    ubijeno3 = []
    globals()["_baza_zelena"] = lazna
    try:
        bu = _BazeUnapred(Zauzeti([r5]), "P", ubij=lambda c: ubijeno3.append(c["id"]),
                          otisak=cist)
        bu.trazi(["S1", "S2"])
        for _ in range(200):
            if "S1" in bu.gotove():
                break
            time.sleep(0.005)
        time.sleep(0.02)
        bu.otkazi()
    finally:
        globals()["_baza_zelena"] = pravi_bz
    tvrdi(not bu._radnici.visi,
          "baza ceka radnika bez odustajanja -- otkaz i Ctrl+C bi cekali da se "
          "neki radnik oslobodi")
    tvrdi(ubijeno3 == [] and "S2" not in bu.gotove(),
          "otkaz ubija Excel radnika koga je baza vec vratila (tudji posao na "
          "istom izlazu), ili pusta sledecu bazu (%r)" % ubijeno3)

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
        bu = _BazeUnapred(_Radnici([r5], korak=0.01), "P",
                          ubij=lambda c: (ubijeno2.append(1), pusti2.set()), otisak=cist)
        bu.trazi(["S1", "S2"])
        t0 = time.time()
        dobijeno = bu.sacekaj({"S1"})
        trajalo = time.time() - t0
    finally:
        globals()["_baza_zelena"] = pravi_bz
    # Granica 0,3 s: gotova baza BUDI onoga ko ceka -- korak cekanja je 0,5 s.
    tvrdi(dobijeno.get("S1") == (True, "ok S1") and ubijeno2 and trajalo < 0.3
          and bu not in _UNAPRED,
          "faza 3 ceka baze koje joj ne trebaju, ili gotova baza ne budi onoga "
          "ko ceka (%.2fs, ubijeno=%r)" % (trajalo, ubijeno2))

    # Bazen: prioritetan zahtev (baza) dobija PRVOG slobodnog radnika, pre posla
    # koji vec ceka; odustajanje vraca None; radnik nikad nije kod dvoje.
    rad = _Radnici([r5], korak=0.01)
    drzi = rad.uzmi()
    redom = []

    def cekaj(ime, prioritet):
        c = rad.uzmi(prioritet=prioritet)
        redom.append(ime)
        time.sleep(0.01)
        rad.vrati(c)

    obican = threading.Thread(target=cekaj, args=("posao", False), daemon=True)
    obican.start()
    time.sleep(0.02)
    baza_nit = threading.Thread(target=cekaj, args=("baza", True), daemon=True)
    baza_nit.start()
    time.sleep(0.02)
    rad.vrati(drzi)
    obican.join(3.0)
    baza_nit.join(3.0)
    tvrdi(redom == ["baza", "posao"],
          "baza ne dobija prvog slobodnog radnika pre posla koji ceka (%r)" % redom)
    # I bez trke niti: dok prioritetan zahtev ceka, obican ne uzima ni
    # slobodnog radnika.
    rad3 = _Radnici([r5], korak=0.01)
    rad3._prioritetnih = 1
    t1 = time.time()
    gotov3, obican_uzeo = u_roku(lambda: rad3.uzmi(
        odustani=lambda: time.time() - t1 > 0.03), rok=1.0)
    rad3._prioritetnih = 0
    tvrdi(gotov3 and obican_uzeo is None
          and u_roku(lambda: rad3.uzmi(prioritet=True), rok=1.0)[1] is r5,
          "obican posao uzima radnika ispred prioritetnog zahteva (baze)")
    rad2 = _Radnici([r5], korak=0.01)
    zauzet = rad2.uzmi()
    t0 = time.time()
    odustao = []
    nit = threading.Thread(target=lambda: odustao.append(
        rad2.uzmi(odustani=lambda: time.time() - t0 > 0.02)), daemon=True)
    nit.start()
    nit.join(1.0)                       # rok: mutant koji ne odustaje ne sme da zaglavi
    tvrdi(odustao == [None] and zauzet is r5,
          "cekanje na radnika ne odustaje (Ctrl+C ili otkaz bi visili)")
    # Poslovi dobijaju radnika PO REDU i kad baza drzi jednog: nit koja vrati
    # radnika ne sme da ga odmah uzme za svoj kasniji posao ispred ranijeg koji
    # ceka. Granica STOP-a stiti samo kasnije poslove -- raniji koji krene
    # posle REVERT-FAIL-a kasnijeg dobio bi njegovu nevracenu kopiju.
    rA, rB = {"id": 1}, {"id": 2}
    rad4 = _Radnici([rA, rB], korak=0.01)
    drzi_bazu = rad4.uzmi(prioritet=True)
    krenuli = []

    def posao_po_redu(k):
        def fn(c):
            krenuli.append(k)
            time.sleep(0.01)
            return k
        return fn

    def po_redu():
        with contextlib.closing(_izvrsi([posao_po_redu(k) for k in range(6)],
                                        rad4)) as tok:
            return list(tok)

    gotov4, rez4 = u_roku(po_redu, rok=3.0)
    rad4.vrati(drzi_bazu)
    tvrdi(gotov4 and krenuli == list(range(6)) and rez4 == list(range(6)),
          "posao krece pre ranijeg koji ceka radnika (raniji bi dobio nevracenu "
          "kopiju kasnijeg posle REVERT-FAIL-a) (%r)" % krenuli)
    tvrdi(_rana_mutacija(4, 5, True) and not _rana_mutacija(5, 5, True)
          and _rana_mutacija(5, 5, False) and not _rana_mutacija(2, 2, True),
          "rana mutacija krece i kad stavke ne staju na radnike koje baza ne drzi")

    # --- ozicenje main-a i _prolaz-a (tekst; main trazi Excel) -----------------
    izvor_main = inspect.getsource(main)
    izvor_prolaz = inspect.getsource(_prolaz)
    tvrdi(izvor_main.count("_izvrsi(") >= 3
          and "_serijska_traka(stavke, args.radnika)" in izvor_main
          and "_zavrsi_istoriju(bazen, args.radnika)" in izvor_main,
          "main ne pusta sve tri faze kroz radnike, ili preskace serijsku traku/istoriju")
    tvrdi(all('izlaz=ctx["izlaz"]' in r or 'izlaz=c["izlaz"]' in r
              for r in izvor_main.splitlines() if "_baza_zelena(" in r),
          "baza cele suite ne ide u izlaz radnika (pisala bi u tests/)")
    tvrdi("    if args.radnika > 1 and izbor:\n" in izvor_main
          and "    radnici = _Radnici(bazen)\n" in izvor_main
          and "            unapred = _BazeUnapred(radnici, pre)\n" in izvor_main
          and "bazen[-1]" not in izvor_main
          and izvor_main.count("radnici)) as tok:") == 3
          and izvor_main.count("posao(g, len(g) > 1, izbor, baze_izbora, True)") == 1
          and izvor_main.count("posao([s], False, izbor, baze_izbora, True)") == 1
          and "return lambda c: najavi(_prolaz(" in izvor_main
          and izvor_main.count("unapred.trazi(") == 1
          and 'kandidati = [s for s, _r in ishod.get("cela_suite", [])]' in izvor_main
          and "ostatak" not in izvor_main[izvor_main.find("    def najavi("):
                                         izvor_main.find("    def posao(")]
          and "        if unapred is not None and isinstance(ishod, dict):\n" in izvor_main
          and '        moguce = ({s["suite"] for s in stavke} & izbor) - set(cele_baze)\n' in izvor_main
          and ("    if unapred is not None and not unapred.suite:\n"
               "        unapred.otkazi()") in izvor_main
          and izvor_main.count("unapred.otkazi()") == 2
          and "nove = unapred.trazi(sorted({s[\"suite\"] for s in kandidati} & moguce))" in izvor_main
          and "unapred.otkazi()" in izvor_main,
          "baze unapred: ne krecu na zahtev (samo stavka koja ide na celu suite, "
          "iz faza 1-2), faze ne idu kroz zajednicki bazen, ili se baze unapred "
          "ne otkazuju")
    tvrdi(("            c = radnici.uzmi(prioritet=True)\n"
           "            try:\n"
           "                ok_opis = _baza_zelena(suite, src=c[\"kopija\"], izlaz=c[\"izlaz\"])\n"
           "            finally:\n"
           "                radnici.vrati(c)\n") in izvor_main,
          "serijska baza faze 3 ne pozajmljuje radnika (gazila bi radnika baze unapred)")
    # Faza 3 ima JEDAN put ocene (_faza3_ocena) za N=1 i N>1; ovde se tvrdi
    # samo ozicenje, ponasanje meri _self_test_faza3.
    tvrdi(izvor_main.count("_faza3_ocena(") == 1
          and izvor_main.count("upisi(ishod[") == 2
          and "_posao_faze3(s, baza_poznata,\n" in izvor_main
          and "            return _baza_poznata(cele_baze, unapred, suite)\n" in izvor_main
          and "        for suite in _baze_faze3_ovde(potrebne_3, unapred, cele_baze):\n" in izvor_main
          and ("                lambda ishod: koren_tokom.add(ishod.get(\"koren\") or koren_pre),\n"
               "                baza_poznata)\n") in izvor_main
          and ("                if not _uzmi_baze_unapred(unapred, potrebne_3, "
               "cele_baze, nacin):\n                    return None\n"
               "                unapred = None\n") in izvor_main
          and 0 <= izvor_main.find("        if not _rana_mutacija(len(sve), len(radnici),\n"
                                   "                              unapred is not None and "
                                   "not unapred.gotovo()):\n"
                                   "            baze(\"cekaj\")\n")
          < izvor_main.find("_posao_faze3(s, baza_poznata,"),
          "faza 3 nije ozicena kroz zajednicku ocenu, ili rana mutacija ne ceka "
          "bazu kad stavke ne staju na radnike bez rezervisanog")
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


def _self_test_faza3() -> list:
    """Faza 3 u oba smera, PONASANJEM (bez Excela): jedna ocena za N=1 i N>1,
    priznanje samo uz zelenu bazu, STOP koji ne ceka, posao koji se ne pusta
    uz poznatu crvenu bazu, i prenos baza unapred."""
    import io
    nalazi = []

    def tvrdi(uslov, opis):
        if not uslov:
            nalazi.append("FAZA 3: " + opis)

    def st(ime, suite):
        return {"ime": ime, "suite": suite}

    a1, b1, a2, b2 = st("a1", "A"), st("b1", "B"), st("a2", "A"), st("b2", "B")
    zelena_a = {"A": (True, "ok A"), "B": (False, "crvena B")}

    def ish(s, stop="", priznato=True):
        return {"stop": stop, "koren": "k-" + s["ime"],
                "ishodi": [(s, priznato, stop or "OK", priznato)]}

    def pusti(sve, ishodi, baze_fn, baza_fn=lambda suite: None):
        upisano, lose, pozivi, koreni = [], [], [], []

        def upisi(isx):
            for s, _p, stanje, _c in isx:
                upisano.append(s["ime"])
                print("%-46s %s" % (s["ime"], stanje))

        def baze(nacin):
            pozivi.append(nacin)
            return baze_fn(nacin)

        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            stop = _faza3_ocena(sve, iter(ishodi), baze, upisi, lose,
                                lambda i: koreni.append(i.get("koren")), baza_fn)
        return stop, buf.getvalue().splitlines(), upisano, lose, pozivi, koreni

    # N=1: baze poznate unapred; crvena stavka se ne pusta (_PRESKOCENA).
    s1 = pusti([a1, b1, a2, b2],
               [ish(a1), dict(_PRESKOCENA), ish(a2), dict(_PRESKOCENA)],
               lambda nacin: zelena_a)
    # N>1: baza stize tek pri prvom ishodu; crvene mutacije su VEC pustene i
    # "priznate" -- ne smeju da se upisu.
    s4 = pusti([a1, b1, a2, b2], [ish(a1), ish(b1), ish(a2), ish(b2)],
               lambda nacin: None if nacin == "proba" else zelena_a)
    tvrdi(s1[0] == "" and s1[2] == ["a1", "a2"]
          and [x[0] for x in s1[3]] == ["b1", "b2"]
          and [x.split()[0] for x in s1[1][:2]] == ["b1", "b2"]
          and all("NIJE IZMERENA" in x for x in s1[1][:2])
          and s1[4] == ["proba"] and s1[5] == ["k-a1", None, "k-a2", None],
          "serijski: crvene stavke se ne javljaju PRVE i redom kataloga, se "
          "upisuju, ili se koren ne belezi (%r)" % (s1[1:],))
    tvrdi(s4[1] == s1[1] and s4[2] == s1[2] and s4[3] == s1[3]
          and s4[4] == ["proba", "cekaj"]
          and s4[5] == ["k-a1", "k-b1", "k-a2", "k-b2"],
          "N>1 i N=1 ne daju isti izlaz faze 3, ili se koren ne belezi "
          "(%r / %r)" % (s4[1], s1[1]))

    # STOP pre baze: ne ceka, stavka NIJE IZMERENA.
    s = pusti([a1, a2], [ish(a1, stop="REVERT-FAIL x", priznato=False)],
              lambda nacin: None if nacin in ("proba", "stop") else zelena_a)
    tvrdi(s[0] == "REVERT-FAIL x" and "cekaj" not in s[4] and s[2] == []
          and s[3] and "dokaz je stao pre nje" in s[3][0][1] and s[5] == ["k-a1"],
          "STOP ceka bazu koja jos traje, ili se stavka ocenjuje bez baze (%r)" % (s,))
    # STOP dok DRUGA potrebna baza traje, a sopstvena je gotova: ocena njom.
    s = pusti([a1, b1], [ish(a1, stop="REVERT-FAIL x", priznato=False)],
              lambda nacin: None if nacin in ("proba", "stop") else zelena_a,
              lambda suite: zelena_a.get(suite) if suite == "A" else None)
    tvrdi(s[0] == "REVERT-FAIL x" and s[2] == ["a1"] and "cekaj" not in s[4]
          and not any("dokaz je stao" in x[1] for x in s[3]),
          "STOP ne uzima sopstvenu gotovu bazu stavke dok druga traje (%r)" % (s,))
    s = pusti([b1, a1], [ish(b1, stop="REVERT-FAIL x", priznato=True)],
              lambda nacin: None if nacin in ("proba", "stop") else zelena_a,
              lambda suite: zelena_a.get(suite) if suite == "B" else None)
    tvrdi(s[2] == [] and [x[0] for x in s[3]] == ["b1"]
          and bool(s[1]) and "nije zelena" in s[1][0],
          "STOP uz sopstvenu CRVENU bazu priznaje mutaciju (%r)" % (s,))
    # STOP uz potrebnu bazu koja je VEC gotova: uzima se, stavka nosi svoj ishod.
    s = pusti([a1, a2], [ish(a1, stop="REVERT-FAIL x", priznato=False)],
              lambda nacin: None if nacin == "proba" else zelena_a)
    tvrdi(s[0] == "REVERT-FAIL x" and s[2] == ["a1"] and s[4] == ["proba", "stop"]
          and not any("dokaz je stao" in x[1] for x in s[3]),
          "STOP odbacuje potrebnu bazu koja je vec gotova (%r)" % (s,))
    # Prekid i baza koje nema.
    s = pusti([a1], [None], lambda nacin: zelena_a)
    tvrdi(s[0] == "PREKID", "prekid faze 3 se ne javlja (%r)" % (s[0],))
    s = pusti([a1], [ish(a1)], lambda nacin: {})
    tvrdi(s[2] == [] and [x[0] for x in s[3]] == ["a1"],
          "mutacija bez baze svoje suite se priznaje (%r)" % (s,))

    # Baza poznata bez cekanja: izmerena ima prednost, gotova unapred se vidi,
    # nepoznata je None.
    class LU:
        suite = ["B", "C"]

        def rezultat(self, suite):
            return {"B": (False, "crvena B")}.get(suite)

    tvrdi(_baza_poznata({}, LU(), "B") == (False, "crvena B")
          and _baza_poznata({}, LU(), "A") is None
          and _baza_poznata({"B": (True, "ok B")}, LU(), "B") == (True, "ok B")
          and _baza_poznata({}, None, "B") is None,
          "posao faze 3 ne vidi bazu koja je gotova unapred (pusta mutaciju uz "
          "poznatu crvenu bazu)")
    tvrdi(_baze_faze3_ovde({"A", "B", "D"}, LU(), {"A": (True, "")}) == ["D"]
          and _baze_faze3_ovde({"A", "B"}, None, {}) == ["A", "B"],
          "faza 3 serijski meri i bazu koju vec meri baza unapred (dve pune "
          "suite uporedo, obe pisu marker)")

    # Posao faze 3: uz poznatu crvenu bazu se ne pusta.
    pusteno = []
    rez = _posao_faze3(b1, lambda suite: (False, "x"),
                       lambda c: pusteno.append(c))("r1")
    rez2 = _posao_faze3(a1, lambda suite: None,
                        lambda c: pusteno.append(c) or "prolaz")("r2")
    rez3 = _posao_faze3(a1, lambda suite: (True, "ok"),
                        lambda c: pusteno.append(c) or "prolaz")("r3")
    tvrdi(rez == _PRESKOCENA and rez is not _PRESKOCENA and rez2 == rez3 == "prolaz"
          and pusteno == ["r2", "r3"],
          "posao faze 3 pusta mutaciju uz poznatu crvenu bazu, ili ne pusta "
          "uz zelenu/nepoznatu (%r)" % (pusteno,))

    # Prenos baza unapred: potrebne gotove -> uzima se i uz STOP; nepotrebna
    # koja jos traje ne zadrzava; "cekaj" ceka.
    class Lazna:
        def __init__(self, gotove, gotovo):
            self.suite, self._g, self._d = ["S1", "S2"], set(gotove), gotovo
            self.cekano = []

        def gotove(self):
            return set(self._g)

        def gotovo(self):
            return self._d

        def sacekaj(self, potrebne):
            self.cekano.append(set(potrebne))
            return {"S1": (True, "ok S1"), "S2": (False, "crvena S2")}

    def uzmi(gotove, gotovo, potrebne, nacin):
        lu, cb = Lazna(gotove, gotovo), {}
        with contextlib.redirect_stdout(io.StringIO()):
            ok = _uzmi_baze_unapred(lu, set(potrebne), cb, nacin)
        return ok, cb, lu.cekano

    r = uzmi({"S1"}, False, {"S1"}, "stop")
    tvrdi(r[0] and r[1] == {"S1": (True, "ok S1")} and r[2] == [{"S1"}],
          "STOP ne uzima potrebnu gotovu bazu dok nepotrebna traje (%r)" % (r,))
    r = uzmi({"S1"}, False, {"S2"}, "stop")
    r2 = uzmi({"S1"}, False, {"S2"}, "proba")
    tvrdi(not r[0] and not r2[0] and r[2] == [] and r2[2] == [] and r[1] == {},
          "baza koja jos traje se ceka bez \"cekaj\" (%r %r)" % (r, r2))
    r = uzmi({"S1"}, False, {"S2"}, "cekaj")
    tvrdi(r[0] and r[1] == {"S2": (False, "crvena S2")} and r[2] == [{"S2"}],
          "\"cekaj\" ne ceka potrebnu bazu (%r)" % (r,))
    r = uzmi(set(), True, {"S2"}, "stop")
    tvrdi(r[0] and r[2] == [{"S2"}],
          "zavrsen posao baza (i izuzetak iz njega) se ne preuzima (%r)" % (r,))
    return nalazi


def _self_test_g1() -> list:
    """G1 (--prijem) u oba smera, PONASANJEM gde se moze: verdikt izbora se
    hvata iz pravog _prolaz-a, poredjenje ne meri celu suite prema samoj sebi,
    imenilac ne broji placebo, uzorak je deterministican i ne bira po prefiksu."""
    nalazi = []

    def tvrdi(uslov, opis):
        if not uslov:
            nalazi.append("G1: " + opis)

    IZB = frozenset({SUITE_ALL})
    ZELENA = "TESTS   RunAllTests: 3 ukupno, 0 palo\n"
    CRVENA = "TESTS   RunAllTests: 3 ukupno, 1 palo\nFAIL T_A -- sam ne prolazi\n"
    a = _s("a", tvrdnja="razlog a")

    def lanac(izlaz, **kw):
        try:
            rez, _p = _lazni_prolaz([a], izlaz, grupno=False, **kw)
        except Exception as e:              # pad je nalaz, ne goli traceback
            nalazi.append("G1: lanac je pukao -- %r" % (e,))
            return {}, {}
        return rez, _prijem_ishod(rez)

    def imena_cele(p):
        return [s["ime"] for s, _r in p.get("cela_suite", [])]

    def prijem(p):
        return [(s["ime"],) + tuple(r) for s, *r in p.get("prijem", [])]

    # --- 1. lanac kroz pravi _prolaz -> _prijem_ishod ---------------------
    _r, p1a = lanac("FAIL T_A -- razlog a\n", izbor=IZB, baze={}, izlaz_baze=ZELENA)
    tvrdi(prijem(p1a) == [("a", True, "OK", "izbor")] and p1a.get("ishodi") == []
          and imena_cele(p1a) == ["a"],
          "1a: priznanje izbora se ne hvata kao 'izbor', ostaje u ishodi, ili "
          "stavka nije TACNO jednom u cela_suite (%r)" % (p1a,))
    _r, p1b = lanac(ZELENA, izbor=IZB, baze={}, izlaz_baze=ZELENA)
    tvrdi(prijem(p1b) == [("a", False, "u izboru: NE OBARA NISTA", "povratak")]
          and imena_cele(p1b) == ["a"] and p1b.get("ishodi") == [],
          "1b: izbor koji ne obara nije 'povratak', ili stavka nije TACNO jednom "
          "u cela_suite (%r)" % (p1b,))
    _r, p1c = lanac("FAIL T_A -- razlog a\n", izbor=IZB, baze={}, izlaz_baze=CRVENA)
    tvrdi([v[3] for v in prijem(p1c)] == ["povratak"] and imena_cele(p1c) == ["a"]
          and not prijem(p1c)[0][2].startswith("u izboru:"),
          "1c: crvena baza izbora nije 'povratak' sa razlogom baze (%r)" % (p1c,))
    _r, p1d = lanac("FAIL T_A -- razlog a\n", izbor=IZB, baze={},
                    izlaz_baze=ZELENA, timeout=True)
    tvrdi(prijem(p1d) == [("a", False, "TIMEOUT suite", "konacno")]
          and imena_cele(p1d) == ["a"] and p1d.get("ishodi") == [],
          "1d: timeout izbora nije 'konacno', ili ne ide i celom suite-om (%r)" % (p1d,))
    tvrdi(all(k in p1a for k in ("koren", "izbor", "ostatak")),
          "1: _prijem_ishod gubi kljuceve koren/izbor/ostatak")
    rez, p1e = lanac(ZELENA)
    tvrdi(p1e is rez and "prijem" not in rez,
          "1e: prolaz BEZ izbora se transformise")
    rez, p1f = lanac("FAIL T_A -- razlog a\n", izbor=IZB, baze={},
                     izlaz_baze=ZELENA, potpis="DRUGI")
    tvrdi(rez.get("stop", "").startswith("REVERT-FAIL") and p1f is rez
          and "prijem" not in rez,
          "1e: prolaz sa stop-om (REVERT-FAIL) se transformise")

    # --- 2. _g1_klasa: svaki slucaj se od ISTO razlikuje po jednoj osi ------
    IZ_OK = (True, "OK", "izbor")
    for izb, cela, ocekivano in (
            (IZ_OK, (True, "OK"), ("ISTO", "")),
            (IZ_OK, (True, "OK (uz jos 2 testa)"), ("ISTO", "")),
            (IZ_OK, (False, "NE OBARA NISTA"), ("OPASNA", "MASKA")),
            (IZ_OK, (False, "PALA DRUGA TVRDNJA: x"), ("OPASNA", "DRUGI PAD")),
            (IZ_OK, (False, "EXCEL UBIJEN -- d"), ("NEJASNO", "")),
            ((False, "u izboru: NE OBARA NISTA", "povratak"), (True, "OK"),
             ("SKUPA", "cela OK")),
            ((False, "u izboru: NE OBARA NISTA", "povratak"),
             (False, "NE OBARA NISTA"), ("SKUPA", "cela NE")),
            ((False, "baza izbora: 1 palo", "povratak"), (True, "OK"),
             ("SKUPA", "baza izbora crvena")),
            ((False, "RunAllTests: TESTS RunAllTests: 3 ukupno, 1 palo", "povratak"),
             (True, "OK"), ("SKUPA", "baza izbora crvena")),
            # Baza koja NIJE odradjena nije S-4a regresija (review G1).
            ((False, "RunAllTests: timeout", "povratak"), (True, "OK"),
             ("SKUPA", "baza izbora infra")),
            ((False, "RunBusinessFlowProSuite: EXCEL UBIJEN -- x", "povratak"),
             (True, "OK"), ("SKUPA", "baza izbora infra")),
            # 'konacno' = izbor nije posmatrao tvrdnju; bez merenja ni na celoj
            # strani par nije ISTO nego NEUPOREDIV (van imenioca, review G1).
            ((False, "APPLY-FAIL -- x", "konacno"), (False, "APPLY-FAIL -- x"),
             ("NEUPOREDIVO", "nijedna strana ne meri")),
            ((False, "TIMEOUT suite", "konacno"), (False, "TIMEOUT suite"),
             ("NEUPOREDIVO", "nijedna strana ne meri")),
            ((False, "EXCEL UBIJEN -- d", "konacno"), (False, "TIMEOUT suite"),
             ("NEUPOREDIVO", "nijedna strana ne meri")),
            ((False, "TIMEOUT suite", "konacno"), (False, "NE OBARA NISTA"),
             ("NEJASNO", "")),
            ((False, "TIMEOUT suite", "konacno"), (True, "OK"), ("NEJASNO", "")),
            ((False, "APPLY-FAIL -- x", "konacno"), (True, "OK"), ("STROZA", "")),
            (None, (True, "OK"), ("NEUPOREDIVO", "")),
            (IZ_OK, None, ("NEUPOREDIVO", ""))):
        tvrdi(_g1_klasa(izb, cela) == ocekivano,
              "2: klasa %r / %r je %r umesto %r"
              % (izb, cela, _g1_klasa(izb, cela), ocekivano))

    # --- 3. lanac bez placeba: zapis izbora + cela suite -> red -------------
    sloj_a = {"a": "slucajni"}
    r3 = _g1_poredi([a], p1a.get("prijem", []), {"a": (False, "NE OBARA NISTA")}, sloj_a)
    tvrdi([(r["ime"], r["klasa"]) for r in r3] == [("a", "OPASNA")]
          and any(l.startswith("G1 OPASNA/MASKA slucajni RAT a |")
                  for l in _g1_ispis(r3, [a], sloj_a)),
          "3: izbor OK + cela NE OBARA NISTA ne daje red OPASNA sa imenom")
    r3 = _g1_poredi([a], p1b.get("prijem", []), {"a": (True, "OK")}, sloj_a)
    tvrdi([(r["ime"], r["klasa"]) for r in r3] == [("a", "SKUPA")],
          "3: povratak + cela OK nije SKUPA (placebo bi usao u poredjenje)")

    # --- 4. _g1_presuda ----------------------------------------------------
    def red(ime, suite=SUITE_ALL, sloj="slucajni", izb=IZ_OK, cela=(True, "OK")):
        k, pt = _g1_klasa(izb, cela)
        return {"ime": ime, "suite": suite, "test": "T_" + ime, "sloj": sloj,
                "klasa": k, "podtip": pt, "izbor": izb, "cela": cela}

    POV = (False, "u izboru: NE OBARA NISTA", "povratak")

    def slucajni(n_rat, n_bfp, **kw):
        return ([red("r%d" % i, **kw) for i in range(n_rat)]
                + [red("b%d" % i, SUITE_BFP, **kw) for i in range(n_bfp)])

    ciljani_skupa = [red("c%d" % i, sloj="ciljani", izb=POV) for i in range(5)]

    def presuda(redovi, stop="", celovit=True):
        return _g1_presuda(redovi, stop, celovit, G1_MIN, G1_MIN_SUITE)

    rc, r = presuda(slucajni(36, 24) + ciljani_skupa)
    tvrdi(rc == 0 and r.startswith("G1 PRIJEM: PROSAO"),
          "4: 36 RAT + 24 BFP presudjenih ISTO ne daje PROSAO (%s)" % r)
    rc, r = presuda(slucajni(30, 20))
    tvrdi(rc == 0 and "5,8%" in r and "nad 50 presudjenih" in r,
          "4: granica za M=50 nije 5,8%% (%s)" % r)
    rc, r = presuda(slucajni(36, 24) + [red("cilj-opasna", sloj="ciljani",
                                            cela=(False, "NE OBARA NISTA"))])
    tvrdi(rc == 1 and "cilj-opasna" in r and "PAO" in r,
          "4: ciljana OPASNA ne obara prijem po imenu (%s)" % r)
    rc, r = presuda(slucajni(36, 24) + [red("stroza", izb=(False, "APPLY-FAIL", "konacno"))])
    tvrdi(rc == 1 and "stroza" in r, "4: STROZA ne obara prijem (%s)" % r)
    rc, r = presuda(slucajni(27, 18) + [red("p%d" % i, izb=POV) for i in range(5)])
    tvrdi(rc == 2, "4: SKUPA se broji u imenilac (45 presudjenih + 5 SKUPA "
                   "daje %d: %s)" % (rc, r))
    rc, r = presuda(slucajni(55, 10))
    tvrdi(rc == 2 and "BFP" in r, "4: 55 RAT + 10 BFP prolazi bez minimuma po suiti (%s)" % r)
    # stop i celovit su odvojene ose (main uz stop daje i celovit=False).
    rc, r = presuda(slucajni(36, 24), stop="REVERT-FAIL: x", celovit=True)
    tvrdi(rc == 2, "4: PROSAO uz STOP")
    rc, r = presuda(slucajni(36, 24), celovit=False)
    tvrdi(rc == 2, "4: PROSAO uz snimak koji se menjao (celovit=False)")
    rc, r = presuda(slucajni(36, 24) + [red("nej", cela=(False, "TIMEOUT suite"))])
    tvrdi(rc == 2 and "nej" in r, "4: NEJASNO stavka ne blokira PROSAO (%s)" % r)
    rc, r = presuda(slucajni(36, 24) + [red("neup", sloj="ciljani", cela=None)])
    tvrdi(rc == 2 and "neup" in r, "4: ciljana NEUPOREDIVO ne blokira PROSAO (%s)" % r)
    rc, r = presuda([red("f%d" % i, sloj="filter") for i in range(60)])
    tvrdi(rc == 2, "4: poziv sa filterom daje PROSAO")
    rc, r = presuda([red("f%d" % i, sloj="filter") for i in range(3)]
                    + [red("fo", sloj="filter", cela=(False, "NE OBARA NISTA"))])
    tvrdi(rc == 1, "4: OPASNA uz filter ne obara prijem")
    rc, r = presuda(slucajni(36, 24) + [red("s%d" % i, izb=POV) for i in range(40)])
    tvrdi(rc == 0 and "SKUPA slucajni 40/100" in r,
          "4: mnogo SKUPA uz dovoljno presudjenih obara prijem (%s)" % r)
    rc, r = presuda(slucajni(36, 24) + ciljani_skupa)
    tvrdi("SKUPA slucajni 0/60, ciljani 5/5" in r,
          "4: udeo SKUPA mesa slojeve -- ciljani je pristrasan (%s)" % r)
    nemereni = [red("t%d" % i, SUITE_BFP, izb=(False, "TIMEOUT suite", "konacno"),
                    cela=(False, "TIMEOUT suite")) for i in range(10)]
    rc, r = presuda(slucajni(25, 15) + nemereni)
    tvrdi(rc == 2 and "nad 40 presudjenih" in r,
          "4: par bez merenja na obe strane ulazi u imenilac (%d: %s)" % (rc, r))
    s4a = G1_S4A_TESTOVI[0]
    ispis = _g1_ispis([dict(red("i"), test=s4a, klasa="SKUPA",
                            podtip="baza izbora infra")], [], {})
    tvrdi(not any("S-4a REGRESIJA" in x for x in ispis),
          "4: pad Excel-a u bazi izbora se javlja kao S-4a REGRESIJA")
    ispis = _g1_ispis([dict(red("c"), test=s4a, klasa="SKUPA",
                            podtip="baza izbora crvena")], [], {})
    tvrdi(("G1 S-4a REGRESIJA: %s" % s4a) in ispis,
          "4: crvena baza izbora S-4a testa nije prijavljena kao REGRESIJA")

    # --- 5. _g1_uzorak nad sintetickim katalogom ---------------------------
    def st(ime, suite, fajl, test):
        return {"ime": ime, "suite": suite, "fajl": fajl, "test": test}

    kat = []
    for suite, pref, n_f in ((SUITE_ALL, "r", 5), (SUITE_BFP, "b", 4),
                             (SUITE_BANKA, "k", 2)):
        for f in range(n_f):
            for t in range(3):
                for j in range(2):
                    kat.append(st("%s%d-%d-%d" % (pref, f, t, j), suite,
                                  "mod%s%d.bas" % (pref, f), "T_%s%d_%d" % (pref, f, t)))
    kat += [st("x", SUITE_ALL, "modX.bas", "T_x"),
            st("x-dva", SUITE_ALL, "modX.bas", "T_x_dva"),
            st("st-r", SUITE_ALL, "modST.bas", "T_st_r"),
            st("st-b", SUITE_BFP, "modST.bas", "T_st_b")]
    S4A, POSLE, PDOK, PSTAT = ("T_r0_0",), ("T_b1_1",), {"x": "NE"}, {"st-r", "st-b", "x"}
    IZBS = {SUITE_ALL, SUITE_BFP}

    def uz(stavke=kat, n=12, s4a=S4A, posle=POSLE, pdok=PDOK, seme="G1"):
        return _g1_uzorak(stavke, n, IZBS, s4a, posle, pdok, PSTAT, seme)

    def sl(sloj, vrsta="slucajni"):
        return {i for i, v in sloj.items() if v == vrsta}

    sloj, problem = uz()
    tvrdi(not problem and sloj, "5: uzorak nad ispravnim katalogom daje problem (%s)" % problem)
    tvrdi(uz(list(reversed(kat[7:])) + kat[:7])[0] == sloj,
          "5: izmesan ulaz daje drugi uzorak (bira se po katalogu, ne po hash-u)")
    po_suiti = {su: [s for s in kat if s["ime"] in sl(sloj) and s["suite"] == su]
                for su in (SUITE_ALL, SUITE_BFP, SUITE_BANKA)}
    okvir = [s for s in kat if s["suite"] in IZBS]
    k_rat = int(round(12 * sum(1 for s in okvir if s["suite"] == SUITE_ALL)
                      / float(len(okvir))))
    tvrdi(len(po_suiti[SUITE_ALL]) == k_rat and len(po_suiti[SUITE_BFP]) == 12 - k_rat,
          "5: kvote po suiti nisu %d/%d (%d/%d)" % (
              k_rat, 12 - k_rat, len(po_suiti[SUITE_ALL]), len(po_suiti[SUITE_BFP])))
    izabrani = [s for s in kat if s["ime"] in sl(sloj)]
    tvrdi(len({s["test"] for s in izabrani}) == len(izabrani),
          "5: slucajni sloj ima dve stavke istog testa")
    # Fajlovi bazena RAT: bez ciljanih i statickih (r0..r4 i modX -- 6), kvota 7.
    fajlovi_rat = {s["fajl"] for s in okvir if s["suite"] == SUITE_ALL
                   and sloj.get(s["ime"]) != "ciljani" and s["ime"] not in PSTAT}
    tvrdi(len({s["fajl"] for s in po_suiti[SUITE_ALL]})
          == min(k_rat, len(fajlovi_rat)) == 6,
          "5: round-robin ne pokriva fajlove pre drugog kruga")
    tvrdi(sl(sloj, "ciljani") == {s["ime"] for s in kat if s["test"] in S4A + POSLE}
          | {"x"},
          "5: ciljani sloj nije tacno testovi sa spiska + kljucevi POZNATI_NALAZI_DOKAZ")
    tvrdi(sloj.get("x-dva") != "ciljani",
          "5: 'x' povlaci 'x-dva' (izbor po prefiksu, ne po imenu)")
    nedozvoljeni = set()
    for i in range(20):
        s_i, _p = uz(seme="s%d" % i)
        nedozvoljeni |= {ime for ime in s_i if ime in ("st-r", "st-b")
                         or ime.startswith("k")}
    tvrdi(not nedozvoljeni, "5: banka ili staticki POZNATI u uzorku: %s"
          % sorted(nedozvoljeni))
    tvrdi("T_nema" in uz(s4a=("T_nema",))[1] and "nema" in uz(pdok={"nema": "NE"})[1],
          "5: nepostojece ciljano ime se tiho preskace")
    tvrdi(sl(uz(n=2)[0]) <= sl(uz(n=3)[0]) and sl(uz(n=12)[0]) <= sl(uz(n=16)[0]),
          "5: uzorak za manji N nije podskup uzorka za veci N")
    tvrdi(sl(uz(seme="drugo")[0]) != sl(sloj),
          "5: drugo seme daje isti slucajni deo")

    # --- 6. main: odbijanje pre citanja, i ozicenje ------------------------
    citano = []
    pravi_mod = _modul_sabotaza

    def ne_citaj():
        # Odbijanje mora doci PRE kataloga; mutant koji ga propusti staje ovde,
        # ne pusta katalog (27 s) ni Excel.
        citano.append(1)
        raise RuntimeError("katalog citan pre odbijanja")
    globals()["_modul_sabotaza"] = ne_citaj
    try:
        with contextlib.redirect_stderr(io.StringIO()), \
                contextlib.redirect_stdout(io.StringIO()):
            rcs = [main(["--prijem", "--grupe", "6"]), main(["--prijem", "--izolacija"]),
                   main(["--prijem", "10"]), main(["--prijem", "0"])]
    except BaseException as e:              # noqa: BLE001 -- pad je nalaz
        rcs = [repr(e)]
    finally:
        globals()["_modul_sabotaza"] = pravi_mod
    tvrdi(rcs == [2, 2, 2, 2] and not citano,
          "6: --prijem uz --grupe 6 / --izolacija / N < G1_MIN / izricito 0 nije "
          "odbijen PRE citanja kataloga (%r, citano %d)" % (rcs, len(citano)))
    izvor_main = inspect.getsource(main)
    najavi_telo = izvor_main[izvor_main.find("    def najavi("):
                             izvor_main.find("kandidati =")]
    upisi_telo = izvor_main[izvor_main.find("    def upisi("):
                            izvor_main.find("    def posao(")]
    tvrdi(izvor_main.count("_prijem_ishod(") == 1 and "_prijem_ishod(" in najavi_telo,
          "6: najavi ne transformise ishod pre kandidata, ili se transformise i "
          "posao faze 3")
    tvrdi(izvor_main.count('prijem_izbor.extend(ishod.get("prijem"') == 2
          and 'verdikt[s["ime"]]' in upisi_telo and "return rc_g1" in izvor_main
          and 0 <= izvor_main.find("_faza3_ocena(") < izvor_main.find("_g1_poredi("),
          "6: main ne knjizi verdikt izbora iz obe faze, upisi ne pamti verdikt, "
          "ili G1 ne poredi posle faze 3")
    tvrdi(izvor_main.count("unapred.trazi(") == 1 and izvor_main.count("upisi(ishod[") == 2
          and izvor_main.count("_faza3_ocena(") == 1
          and izvor_main.count("izborom_priznatih += ") == 2,
          "6: postojece igle R-4 / faze 3 / VERDIKT su promenjene")
    return nalazi


def _self_test(tiho: bool = False) -> int:
    nalazi = []
    _STOP.clear()
    nalazi.extend(_self_test_r4())
    nalazi.extend(_self_test_faza3())
    nalazi.extend(_self_test_g1())

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
