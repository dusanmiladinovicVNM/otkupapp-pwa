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
import contextlib
import hashlib
import io
import importlib.util
import inspect
import os
import re
import shutil
import subprocess
import sys
import tempfile

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


def _modul_sabotaza():
    put = os.path.join(ROOT, "tools", "sabotaza.py")
    spec = importlib.util.spec_from_file_location("_sab_za_dokaz", put)
    modul = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modul)
    return modul


def _suite_izbor() -> set:
    """Suite koje primaju `run_vba --test` -- iz SUITES (jedini izvor istine).

    Cita se iz kataloga, ne prepisuje: druga lista bi bila druga stvar koja
    moze da se razidje sa ulaznim tackama u VBA (v. vba_gate, nalaz IZBOR).
    """
    put = os.path.join(ROOT, "tools", "run_vba.py")
    spec = importlib.util.spec_from_file_location("_run_vba_za_dokaz", put)
    modul = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modul)
    return {ime for ime, meta in modul.SUITES.items() if meta.get("izbor")}


def _run_suite(suite: str, testovi=None, bez_markera: bool = False,
               src: str = None) -> tuple:
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


def _pripremi_radni(sab, koren_src: str = None) -> tuple:
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
           "alat": os.path.join(radni, "alat", "sabotaza.py")}
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
    shutil.copytree(ctx["snimak"], ctx["kopija"])
    if not _van_radnog_stabla(ctx["kopija"], ctx):
        return "STRAZA: kopija za mutacije je u radnom stablu ili u snimku"
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


def _baza_zelena(suite: str, testovi=None, src: str = None) -> tuple:
    """(je_zelena, opis). Tvrdi se, ne gleda se.

    Uz `testovi` je baza samo taj izbor -- tacno ono sto ce prolaz sa mutacijom
    pustiti. Izbor sa nula provera nije zelen nego prazan: nista nije mereno.
    `src` je kopija nad kojom se meri (D-6); izmedju prolaza je jednaka snimku.
    """
    try:
        r = _pusti(*_run_suite(suite, testovi, src=src))
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
        if kljuc not in baze:
            baze[kljuc] = _baza_zelena(suite, testovi, src=ctx["kopija"])
        ok, opis = baze[kljuc]
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
        run = _pusti(*_run_suite(suite, testovi, bez_markera=True,
                                 src=ctx["kopija"]))
        izlaz = run.stdout + run.stderr
        fatal = _fatal(izlaz)
        if fatal:
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


def _izolacija(stavke: list, izbor: set, src: str = None) -> list:
    """[(suite, test, opis)] za ciljane testove koji SAMI ne prolaze.

    Bez mutacija: svaki razlicit ciljani test se pusti sam nad cistim izvorom.
    Crven test nije kvar izbora nego nalaz o testu -- zavisi od stanja koje mu
    ostavlja raniji test u suite-u. `--izbor` ga ionako meri celom suite-om, a
    ovde se nabraja da se zna koliko testova tako zavisi od redosleda.
    """
    crveni = []
    vidjeni = set()
    for s in stavke:
        kljuc = (s["suite"], s["test"])
        if s["suite"] not in izbor or kljuc in vidjeni:
            continue
        vidjeni.add(kljuc)
        ok, opis = _baza_zelena(s["suite"], [s["test"]], src=src)
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

    # --- D-6: snimak i kopija PRE prve baze --------------------------------
    #
    # Sve od ovog trenutka meri SNIMAK: baze i mutacije idu nad kopijom, a
    # src-vba se samo cita (potpis pre/tokom/posle). Procedure sidara se citaju
    # iz snimka, da plan grupa i mutacija gledaju isti tekst.
    global _RADNI
    ctx, problem = _pripremi_radni(sab)
    _RADNI = ctx["radni"]
    if problem:
        print("STOP: " + problem, file=sys.stderr)
        shutil.rmtree(ctx["radni"], ignore_errors=True)
        return 2
    stavke = _izaberi(sab, args.filter, ctx["snimak"])

    izbor = _suite_izbor() if (args.izbor or args.izolacija) else set()
    if args.izolacija:
        try:
            return 1 if _izolacija(stavke, izbor, ctx["kopija"]) else 0
        finally:
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
        ok, opis = _baza_zelena(suite, src=ctx["kopija"])
        cele_baze[suite] = (ok, opis)
        print("BAZNO: %s" % opis, flush=True)
        if not ok:
            print("STOP: baza nije zelena. Dokaz bi merio crveno koje sabotaza "
                  "nije izazvala. Kopija ostaje u %s" % ctx["radni"],
                  file=sys.stderr)
            return 2

    pre = _otisak(ctx["snimak"])
    # src-vba je bio jednak snimku u trenutku kopiranja (tvrdi _pripremi_radni);
    # izmena tokom baza zato mora da se vidi -- ne uzima se nov uzorak.
    koren_pre = pre
    koren_tokom = set()
    print("potpis izvora: %s (snimak; src-vba se ne pise)" % pre, flush=True)

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

    for i, g in enumerate(grupe, 1):
        if len(g) > 1:
            print("GRUPA %d/%d  %s  clanova %d"
                  % (i, len(grupe), g[0]["suite"], len(g)), flush=True)
        ishod = _prolaz(g, pre, len(g) > 1, izbor, baze_izbora, ctx)
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
            break

    if ponavljaju and not stop:
        print("PONAVLJA SE SAM: %d" % len(ponavljaju), flush=True)
        for s in ponavljaju:
            ishod = _prolaz([s], pre, False, izbor, baze_izbora, ctx)
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
                break

    # Test koji SAM ne prolazi meri se celom suite-om -- uz bazu CELE suite,
    # koja se tvrdi jednom po suite-u i tek kad zatreba.
    if cela_suite and not stop:
        print("MERI SE CELOM SUITE-OM: %d (ciljani test sam ne prolazi, ili izbor "
              "nije oborio tvrdnju -- v. --izolacija)" % len(cela_suite), flush=True)
        for s, razlog in cela_suite:
            print("  %-46s <- %s" % (s["ime"], razlog), flush=True)
            if s["suite"] not in cele_baze:
                cele_baze[s["suite"]] = _baza_zelena(s["suite"], src=ctx["kopija"])
                print("BAZNO: %s" % cele_baze[s["suite"]][1], flush=True)
            ok, opis = cele_baze[s["suite"]]
            if not ok:
                lose.append((s["ime"], "NIJE IZMERENA: ni izbor ni cela suite "
                                       "nisu zeleni -- %s" % opis))
                print("%-46s NIJE IZMERENA: baza cele suite nije zelena"
                      % s["ime"], flush=True)
                continue
            ishod = _prolaz([s], pre, False, ctx=ctx)
            koren_tokom.add(ishod.get("koren") or koren_pre)
            upisi(ishod["ishodi"])
            if ishod["stop"]:
                stop = ishod["stop"]
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
    if stop:
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
    tvrdi("ctx, problem = _pripremi_radni(sab)" in izvor_main,
          "D-6: main ne pravi snimak i kopiju pre prve baze")
    tvrdi("koren_pre = pre\n" in izvor_main
          and "_koren_menjan(pre, koren_tokom, _otisak())" in izvor_main,
          "D-6: main poredi radno stablo sa kasnijim uzorkom umesto sa snimkom")
    for ime, igla in (("_prolaz(", "ctx"), ("_baza_zelena(", 'src=ctx["kopija"]'),
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
        print("\nprekinuto -- src-vba nije diran (dokaz radi nad kopijom%s)"
              % (" u " + _RADNI if _RADNI else ""), file=sys.stderr)
        sys.exit(130)
