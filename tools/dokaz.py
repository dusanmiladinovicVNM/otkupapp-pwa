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

SESIJA NE CEKA DOKAZ. Nikad. Run se pusta odvojeno i verdikt se cita iz fajla:

    powershell -File tools\dokaz_bg.ps1 modOtkupUI.bas   # pusti i vrati se ODMAH
    python tools/dokaz.py --status                       # gde je stao / verdikt
    python tools/dokaz.py --knjiga modOtkupUI.bas        # preskoci nepromenjeno

Verdikt se upisuje u `tests/dokaz_last.json` POSLE SVAKE SABOTAZE, ne tek na
kraju -- zato `--status` odgovara trenutno i sredinom run-a ("18/38, sve OK") i
niko nema razlog da blokira na cekanju. Pun katalog jednom nocno radi
`tools\dokaz_nocu.ps1` (Scheduled Task; ujutru stoji `tests/dokaz_jutro.md`).

TRI STVARI KOJE OVAJ ALAT TVRDI, a ne samo gleda:

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
import datetime
import hashlib
import importlib.util
import io
import json
import os
import re
import shutil
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_VBA = os.path.join(ROOT, "src-vba")
TESTS_DIR = os.path.join(ROOT, "tests")
SUITE_BANKA = "RunBankaImportTestSuite"
SUITE_ALL = "RunAllTests"

# Verdikt i knjiga zive pored ostalih rezultata run-a (tests/last_run.json), i
# kao i oni su gitignored: oba zavise od fixture-a i masine, pa u repou ne bi
# znacila isto.
JSON_DEFAULT = os.path.join(TESTS_DIR, "dokaz_last.json")
KNJIGA = os.path.join(TESTS_DIR, "dokaz_ledger.json")
KNJIGA_VERZIJA = 1
# Sva tri ulaze u kljuc knjige: popravka u bilo kom od njih menja sta dokaz
# uopste meri, pa svi stari unosi moraju da padnu.
ALATI = ("dokaz.py", "sabotaza.py", "run_vba.py")


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


def _suite_za(test: str) -> str:
    # Pisac ide u banka-suite: RunAllTests je nemutirajuca, pa tvrdnji o upisu
    # u njoj nema.
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
    return _tokeni_banke(izlaz)


def _kljuc_testa(test: str, suite: str) -> str:
    """Sta se poredi sa onim sto je palo."""
    if suite == SUITE_ALL:
        return test
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



# --- verdikt u fajl, ne u ocekivanje -----------------------------------------
#
# Dokaz nad punim katalogom traje satima. Dok je jedini izlaz bio stdout, neko
# (operater ili agent) je morao da SEDI nad njim: sesija je blokirala u
# desetominutnim blokovima, trosila kontekst i kvotu na cekanje, a verdikt je
# posle svega postojao samo u scrollback-u.
#
# Zato svaki run pise masinski citljiv verdikt POSLE SVAKE SABOTAZE, ne tek na
# kraju. Time `--status` odgovara trenutno i sredinom run-a ("18/38, sve OK"),
# pa niko nema razlog da blokira. Fajl se pise i kad run pukne -- prekinut dokaz
# je nalaz koji se sutra cita, ne rupa.
def _grana_i_commit() -> tuple:
    def g(*a):
        try:
            r = subprocess.run(a, cwd=ROOT, capture_output=True, text=True,
                               timeout=15)
            return r.stdout.strip() if r.returncode == 0 else ""
        except Exception:                        # noqa: BLE001
            return ""
    return g("git", "rev-parse", "--abbrev-ref", "HEAD"), g("git", "rev-parse", "HEAD")


def _upisi_json(put: str, podaci: dict) -> None:
    if not put:
        return
    try:
        mapa = os.path.dirname(os.path.abspath(put))
        if mapa:
            os.makedirs(mapa, exist_ok=True)
        with open(put, "w", encoding="utf-8") as fh:
            json.dump(podaci, fh, ensure_ascii=True, indent=2)
    except OSError as e:
        print("UPOZORENJE: verdikt nije upisan u %s -- %s" % (put, e),
              file=sys.stderr)


def status(put: str) -> int:
    """Procitaj poslednji verdikt. Izlaz: 0 dokazano, 1 nije, 2 nema, 3 u toku."""
    try:
        with open(put, "r", encoding="utf-8") as fh:
            d = json.load(fh)
    except (OSError, ValueError) as e:
        print("nema verdikta u %s (%s)" % (put, e), file=sys.stderr)
        return 2

    print("%s   grana=%s   pocet=%s" % (d.get("verdikt", "?"),
                                        d.get("grana", "?"),
                                        d.get("pocet", "?")))
    print("sabotaza %s: obradjeno %s, crvenih %s, preneseno %s, problema %s"
          % (d.get("sabotaza"), d.get("obradjeno"), d.get("crvenih"),
             d.get("preneseno"), len(d.get("problemi") or [])))
    for ime, sta in (d.get("problemi") or []):
        print(" PROBLEM: %s -> %s" % (ime, sta))
    if d.get("u_toku"):
        # U TOKU nije ni zeleno ni crveno. Poseban izlaz, da ga skripta ne
        # pomesa sa verdiktom.
        print("run JOS TRAJE (ili je prekinut bez zavrsetka).")
        return 3
    return 0 if d.get("verdikt") == "DOKAZANO" else 1


# --- knjiga dokazanog --------------------------------------------------------
#
# Sabotaza dokazuje da test NIJE placebo. To je cinjenica o paru (test, kod
# ispod njega) i ne zastareva dok se taj par ne promeni. Ponovno dokazivanje
# celog kataloga na kraju svakog kruga ispravki ne meri nista novo, a kosta sat.
#
# ZASTO OVO NE SLABI KAPIJU -- tri ograde, sve tri nose svoj deo:
#
#   1. Knjiga se cita SAMO na izricit `--knjiga` i SAMO uz filter. Run bez
#      filtera (dakle i nocni) je uvek pun. Zato je zastarelost ogranicena na
#      jednu noc -- jaca garancija nego bilo koja heuristika nad starosti unosa.
#   2. Verdikt UVEK kaze koliko je preneseno, a kad nista nije mereno to pise u
#      samoj liniji verdikta. "DOKAZANO" ne moze da se procita kao svez dokaz.
#   3. Ostecena knjiga, drugacija verzija, ili kljuc koji se ne poklapa -> unos
#      se ignorise i dokazuje se ponovo. Sumnja uvek ide na skuplju stranu.
#
# Kljuc namerno pokriva vise nego sto sabotaza dira: sam unos u katalogu (svih
# pet clanova), mutirani fajl, modul u kom test zivi, sva tri alata i potpis
# fixture-a. Kad se modul testa ne nadje, kljuc pada na potpis CELOG src-vba --
# konzervativno (bilo koja izmena bilo gde ponistava unos), nikad obrnuto.
def _hash_bajtova(*delovi) -> str:
    h = hashlib.sha256()
    for d in delovi:
        h.update(d if isinstance(d, bytes) else repr(d).encode("utf-8"))
        h.update(b"\x00")                        # granica, da se delovi ne sliju
    return h.hexdigest()[:16]


def _hash_fajla(put: str) -> str:
    try:
        with open(put, "rb") as fh:
            return hashlib.sha256(fh.read()).hexdigest()[:16]
    except OSError:
        return "?"


def _hash_alata() -> str:
    return _hash_bajtova(*[_hash_fajla(os.path.join(ROOT, "tools", a))
                           for a in ALATI])


def _potpis_fixture() -> str:
    """Potpis podataka fixture-a. "?" kad se generator ne ucita -- nije kapija,
    samo ulazi u kljuc, pa nepoznata vrednost cini kljuc stabilnim na svoj nacin.
    """
    put = os.path.join(ROOT, "tools", "make_fixture.py")
    try:
        spec = importlib.util.spec_from_file_location("_mf_za_dokaz", put)
        modul = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(modul)
        return modul.signature()
    except Exception:                            # noqa: BLE001
        return "?"


_DEF_TESTA = None


def _modul_testa(test: str) -> str:
    """Fajl u src-vba koji definise tu proceduru; "" kad se ne nadje.

    Prazan rezultat NIJE greska nego signal da kljuc mora da padne na ceo
    src-vba: ime iz banka-suite (T21_...) se ne poklapa sa imenom procedure
    koja se stvarno izvrsava, pa se za takve unose ne sme praviti uzak kljuc.
    """
    global _DEF_TESTA
    if _DEF_TESTA is None:
        _DEF_TESTA = {}
        wzor = re.compile(
            r"^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*"
            r"(?:Sub|Function)\s+([A-Za-z0-9_]+)", re.I)
        for ime in sorted(os.listdir(SRC_VBA)):
            if os.path.splitext(ime)[1].lower() not in (
                    ".bas", ".cls", ".frm", ".doccls"):
                continue
            put = os.path.join(SRC_VBA, ime)
            try:
                with open(put, "rb") as fh:
                    tekst = fh.read().decode("ascii", errors="replace")
            except OSError:
                continue
            for red in tekst.split("\n"):
                m = wzor.match(red)
                if m:
                    _DEF_TESTA.setdefault(m.group(1).lower(), put)
    return _DEF_TESTA.get((test or "").lower(), "")


def _kljuc_unosa(unos: tuple, hash_alata: str, sig_fix: str,
                 otisak_svega: str) -> str:
    fajl, staro, novo, test, tvrdnja = unos
    modul = _modul_testa(test)
    hash_testa = _hash_fajla(modul) if modul else otisak_svega
    return _hash_bajtova(fajl, staro, novo, test, tvrdnja,
                         _hash_fajla(os.path.join(SRC_VBA, fajl)),
                         hash_testa, hash_alata, sig_fix)


def _ucitaj_knjigu() -> dict:
    """Prazna knjiga na svaku sumnju -- osteceno ili starije znaci pun dokaz."""
    try:
        with open(KNJIGA, "r", encoding="utf-8") as fh:
            k = json.load(fh)
    except (OSError, ValueError):
        return {}
    if not isinstance(k, dict) or k.get("verzija") != KNJIGA_VERZIJA:
        return {}
    unosi = k.get("unosi")
    return unosi if isinstance(unosi, dict) else {}


def _upisi_knjigu(unosi: dict) -> None:
    try:
        os.makedirs(TESTS_DIR, exist_ok=True)
        with open(KNJIGA, "w", encoding="utf-8") as fh:
            json.dump({"verzija": KNJIGA_VERZIJA, "unosi": unosi}, fh,
                      ensure_ascii=True, indent=2, sort_keys=True)
    except OSError as e:
        print("UPOZORENJE: knjiga nije upisana -- %s" % e, file=sys.stderr)


# --- dokaz nad samom knjigom -------------------------------------------------
#
# Knjiga preskace dokaze. Provera koja to dozvoljava mora sama biti dokazana --
# inace se, tacno po CLAUDE.md par. 5, jedan sat cekanja zameni za tihu rupu.
# obe polovine: pravila nad ISPRAVNIM ulazom moraju da prodju, i svako pravilo
# kad se PREKRSI mora da se javi BAS PO SVOM IMENU.
#
# Sve radi bez Excela, pa ide i u CI (staticka kapija), za razliku od samog
# dvosmernog dokaza koji trazi Windows.
def _st_slucajevi(kljuc_fn):
    """Vrati listu (ime, poruka_ili_None). None = pravilo drzi."""
    nalazi = []

    def tvrdi(ime, uslov, poruka):
        nalazi.append((ime, None if uslov else poruka))

    basovi = sorted(n for n in os.listdir(SRC_VBA) if n.endswith(".bas"))
    fajlA, fajlB = basovi[0], basovi[1]
    # Ime procedure koja stvarno postoji, da se grana "modul nadjen" izvrsi.
    poznat_test = next((t for t in sorted(_DEF_TESTA or {}) if t.startswith("t_")),
                       "") or sorted(_DEF_TESTA or {"x": 1})[0]

    baza = (fajlA, "staro", "novo", poznat_test, "tvrdnja")
    k = kljuc_fn(baza, "alatX", "fixX", "otisakX")

    tvrdi("kljuc-stabilan",
          k == kljuc_fn(baza, "alatX", "fixX", "otisakX"),
          "isti ulaz daje dva razlicita kljuca")

    for i, polje in enumerate(("fajl", "staro", "novo", "test", "tvrdnja")):
        drugi = list(baza)
        drugi[i] = fajlB if i == 0 else (baza[i] + "-drugo")
        tvrdi("kljuc-unos-" + polje,
              kljuc_fn(tuple(drugi), "alatX", "fixX", "otisakX") != k,
              "izmena '%s' u katalogu ne menja kljuc" % polje)

    tvrdi("kljuc-alat",
          kljuc_fn(baza, "alatY", "fixX", "otisakX") != k,
          "izmena alata ne menja kljuc -- popravljen dokaz bi vazio za stari")
    tvrdi("kljuc-fixture",
          kljuc_fn(baza, "alatX", "fixY", "otisakX") != k,
          "izmena fixture-a ne menja kljuc")

    # Nepoznat test -> kljuc MORA da padne na potpis celog src-vba. Da to zaista
    # radi, vidi se tako sto isti unos sa razlicitim otiskom daje razlicit kljuc.
    nepoznat = (fajlA, "staro", "novo", "T_OvogaNemaNigde_123", "tvrdnja")
    tvrdi("kljuc-nepoznat-test-pada-na-ceo-izvor",
          (kljuc_fn(nepoznat, "alatX", "fixX", "otisakX")
           != kljuc_fn(nepoznat, "alatX", "fixX", "otisakY")),
          "za nepoznat test kljuc ne zavisi od celog src-vba -- unos bi "
          "prezivljavao izmene bilo gde")

    # SADRZAJ fajla, ne njegovo IME.
    #
    # Prva verzija je ovo merila kroz `kljuc-unos-fajl` (dva razlicita imena
    # fajla), pa je prolazila i kad kljuc uopste ne gleda sadrzaj -- ime je
    # ionako i samo clan kljuca. Bas to je i cela svrha kljuca: izmena u
    # mutiranom fajlu mora da ponisti unos. Nadjeno sabotazom nad ovim
    # self-testom, ne pregledom.
    #
    # Test je namerno NEPOZNAT, da `_hash_fajla` u ovom racunu hrani jedino
    # mutirani fajl (za poznat test hrani i modul testa, pa bi provera mogla da
    # prodje na pogresnoj polovini).
    global _hash_fajla
    pravi_hash = _hash_fajla
    try:
        _hash_fajla = lambda _p: "sadrzajA"      # noqa: E731
        kA = kljuc_fn(nepoznat, "alatX", "fixX", "otisakX")
        _hash_fajla = lambda _p: "sadrzajB"      # noqa: E731
        kB = kljuc_fn(nepoznat, "alatX", "fixX", "otisakX")
    finally:
        _hash_fajla = pravi_hash
    tvrdi("kljuc-sadrzaj-fajla", kA != kB,
          "kljuc ne gleda SADRZAJ mutiranog fajla, samo ime -- izmena u kodu "
          "ne bi ponistila unos u knjizi")
    return nalazi


def _st_stvarni_hashevi():
    """Da `_hash_alata` zaista cita alate.

    `_st_slucajevi` prima hash alata kao ARGUMENT, pa dokazuje samo da parametar
    ulazi u kljuc -- a ne i da ga iko stvarno racuna. Sabotaza koja je
    `_hash_alata` zamenila konstantom prosla je neprimeceno; ovo je popravka.
    """
    nalazi = []

    def tvrdi(ime, uslov, poruka):
        nalazi.append((ime, None if uslov else poruka))

    tvrdi("alati-sva-tri",
          set(ALATI) == {"dokaz.py", "sabotaza.py", "run_vba.py"},
          "ALATI vise ne pokriva sva tri alata -- popravka u izostavljenom "
          "ne bi ponistila knjigu")

    global _hash_fajla
    pravi_hash = _hash_fajla
    try:
        _hash_fajla = lambda p: os.path.basename(p)   # noqa: E731
        h1 = _hash_alata()
        _hash_fajla = lambda _p: "isto"               # noqa: E731
        h2 = _hash_alata()
    finally:
        _hash_fajla = pravi_hash
    tvrdi("alati-hash-cita-fajlove", h1 != h2,
          "_hash_alata ne zavisi od sadrzaja alata -- vraca isto ma sta bilo u "
          "dokaz.py/sabotaza.py/run_vba.py")
    return nalazi


def _st_knjiga_i_verdikt():
    """Provere nad ucitavanjem knjige i citanjem verdikta."""
    import contextlib
    import tempfile
    nalazi = []

    def tiho(fn, *a):
        # status() namerno pise coveku; u self-testu se meri samo izlazni kod.
        with contextlib.redirect_stdout(io.StringIO()), \
                contextlib.redirect_stderr(io.StringIO()):
            return fn(*a)

    def tvrdi(ime, uslov, poruka):
        nalazi.append((ime, None if uslov else poruka))

    global KNJIGA
    stara = KNJIGA
    tmp = tempfile.mkdtemp(prefix="dokaz_st_")
    try:
        KNJIGA = os.path.join(tmp, "knjiga.json")

        _upisi_knjigu({"a": {"kljuc": "k1", "verdikt": "OK"}})
        tvrdi("knjiga-krug", _ucitaj_knjigu().get("a", {}).get("kljuc") == "k1",
              "upisan unos se ne procita nazad")

        with open(KNJIGA, "w", encoding="utf-8") as fh:
            fh.write("{ ovo nije json")
        tvrdi("knjiga-osteceno", _ucitaj_knjigu() == {},
              "ostecena knjiga se ne odbacuje -- dokaz bi se preskakao po smecu")

        with open(KNJIGA, "w", encoding="utf-8") as fh:
            json.dump({"verzija": KNJIGA_VERZIJA + 1,
                       "unosi": {"a": {"kljuc": "k1", "verdikt": "OK"}}}, fh)
        tvrdi("knjiga-verzija", _ucitaj_knjigu() == {},
              "knjiga druge verzije se cita -- format bi tiho promenio znacenje")

        os.remove(KNJIGA)
        tvrdi("knjiga-nema", _ucitaj_knjigu() == {},
              "nepostojeca knjiga ne daje praznu")

        vput = os.path.join(tmp, "v.json")
        tvrdi("verdikt-nema", tiho(status, vput) == 2, "nepostojeci verdikt nije 2")
        _upisi_json(vput, {"verdikt": "U TOKU", "u_toku": True})
        tvrdi("verdikt-u-toku", tiho(status, vput) == 3,
              "run u toku se ne razlikuje od zavrsenog")
        _upisi_json(vput, {"verdikt": "NIJE DOKAZANO", "u_toku": False})
        tvrdi("verdikt-crveno", tiho(status, vput) == 1, "crven verdikt nije 1")
        _upisi_json(vput, {"verdikt": "DOKAZANO", "u_toku": False})
        tvrdi("verdikt-zeleno", tiho(status, vput) == 0, "zelen verdikt nije 0")
    finally:
        KNJIGA = stara
        shutil.rmtree(tmp, ignore_errors=True)
    return nalazi


def self_test() -> int:
    _modul_testa("")                             # napuni indeks procedura
    nalazi = (_st_slucajevi(_kljuc_unosa) + _st_stvarni_hashevi()
              + _st_knjiga_i_verdikt())
    pali = [(i, p) for i, p in nalazi if p]
    if pali:
        for ime, poruka in pali:
            print("SELF-TEST: %s -- %s" % (ime, poruka), file=sys.stderr)
        return 2

    # --- druga polovina: svako pravilo mora da UME da pukne ----------------
    #
    # Podmetnuta je pokvarena verzija racunanja kljuca i tvrdi se da nalaz
    # stigne BAS po imenu koje to pravilo nosi. Zelena provera koja nikad nije
    # pokazana crvena ne dokazuje da ista meri.
    def slep_na_tvrdnju(unos, ha, sf, ots):
        return _kljuc_unosa(unos[:4] + ("",), ha, sf, ots)

    def slep_na_alat(unos, ha, sf, ots):
        return _kljuc_unosa(unos, "", sf, ots)

    def slep_na_izvor(unos, ha, sf, ots):
        return _kljuc_unosa(unos, ha, sf, "")

    def slep_na_sadrzaj(unos, ha, sf, ots):
        # Kljuc gleda IME fajla ali ne i njegov sadrzaj -- rupa koju je prva
        # verzija ovog self-testa propustila.
        fajl, staro, novo, test, tvrdnja = unos
        modul = _modul_testa(test)
        return _hash_bajtova(fajl, staro, novo, test, tvrdnja,
                             _hash_fajla(modul) if modul else ots, ha, sf)

    sabotaze = [
        (slep_na_tvrdnju, "kljuc-unos-tvrdnja"),
        (slep_na_alat, "kljuc-alat"),
        (slep_na_izvor, "kljuc-nepoznat-test-pada-na-ceo-izvor"),
        (slep_na_sadrzaj, "kljuc-sadrzaj-fajla"),
    ]
    for fn, ocekivano in sabotaze:
        crveni = {i for i, p in _st_slucajevi(fn) if p}
        if ocekivano not in crveni:
            print("SELF-TEST: sabotaza '%s' NIJE oborila to pravilo (crveni: %s)"
                  % (ocekivano, ", ".join(sorted(crveni)) or "nijedan"),
                  file=sys.stderr)
            return 2

    print("self-test: cisto (%d pravila + %d sabotaza sa dvosmernim dokazom)."
          % (len(nalazi), len(sabotaze)))
    return 0


def main(argv: list) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("filter", nargs="*",
                    help="ime fajla (modX.bas) ili prefiks imena sabotaze")
    ap.add_argument("--json", dest="json_put", nargs="?", const=JSON_DEFAULT,
                    default=JSON_DEFAULT,
                    help="gde ide masinski verdikt (podrazumevano %s)"
                         % os.path.relpath(JSON_DEFAULT, ROOT))
    ap.add_argument("--bez-json", action="store_true",
                    help="ne pisi verdikt u fajl")
    ap.add_argument("--status", action="store_true",
                    help="procitaj poslednji verdikt i izadji (ne pusta nista)")
    ap.add_argument("--knjiga", action="store_true",
                    help="preskoci sabotaze cij se kljuc nije promenio "
                         "(radi SAMO uz filter; pun katalog je uvek pun)")
    ap.add_argument("--self-test", action="store_true",
                    help="dokazi da knjiga i verdikt zaista mere (bez Excela)")
    args = ap.parse_args(argv)

    if args.self_test:
        return self_test()
    if args.status:
        return status(args.json_put)

    json_put = "" if args.bez_json else args.json_put

    sab = _modul_sabotaza()
    katalog = sab.SABOTAZE
    poznati_spisak = getattr(sab, "POZNATI_NALAZI_DOKAZ", {})
    stavke = []
    for ime, unos in katalog.items():
        fajl, _staro, _novo, test, tvrdnja = unos
        if args.filter and not (fajl in args.filter or
                                any(ime.startswith(f) for f in args.filter)):
            continue
        stavke.append((ime, fajl, test, tvrdnja, unos))

    if not stavke:
        print("filter ne pogadja nijednu sabotazu", file=sys.stderr)
        return 2

    grana, commit = _grana_i_commit()
    rezultat = {
        "verdikt": "U TOKU",
        "u_toku": True,
        "pocet": datetime.datetime.now().isoformat(timespec="seconds"),
        "zavrsen": None,
        "sekundi": 0,
        "grana": grana,
        "commit": commit,
        "filter": list(args.filter),
        "knjiga": bool(args.knjiga),
        "sabotaza": len(stavke),
        "obradjeno": 0,
        "crvenih": 0,
        "preneseno": 0,
        "poznatih": 0,
        "stavke": [],
        "problemi": [],
        "poznati": [],
    }
    t0 = time.time()

    def snimi():
        rezultat["sekundi"] = round(time.time() - t0, 1)
        _upisi_json(json_put, rezultat)

    # --- knjiga: sta uopste treba dokazivati ------------------------------
    #
    # Kljucevi se racunaju PRE prve mutacije, dok je izvor zdrav -- posle bi
    # merili mutirano stanje.
    knjiga = _ucitaj_knjigu()
    hash_alata = _hash_alata()
    sig_fix = _potpis_fixture()
    otisak_svega = _otisak()
    kljucevi = {ime: _kljuc_unosa(unos, hash_alata, sig_fix, otisak_svega)
                for ime, _f, _t, _tv, unos in stavke}

    preneseni, za_dokaz = [], []
    # Pun katalog NAMERNO ne cita knjigu. To je ograda zbog koje je knjiga
    # bezbedna: nocni run bez filtera svake noci dokazuje sve iz pocetka.
    koristi_knjigu = args.knjiga and bool(args.filter)
    if args.knjiga and not args.filter:
        print("--knjiga se ignorise bez filtera: pun katalog je uvek pun.",
              flush=True)
    for s in stavke:
        ime = s[0]
        u = knjiga.get(ime) if koristi_knjigu else None
        if u and u.get("kljuc") == kljucevi[ime] and u.get("verdikt") == "OK":
            preneseni.append((ime, u))
        else:
            za_dokaz.append(s)

    rezultat["preneseno"] = len(preneseni)
    for ime, u in preneseni:
        print("%-46s PRENESENO (dokazano %s)" % (ime, u.get("vreme", "?")),
              flush=True)
        rezultat["stavke"].append({"ime": ime, "stanje": "PRENESENO",
                                   "ok": True, "mereno": False})

    print("sabotaza: %d (za dokaz %d, preneseno %d)"
          % (len(stavke), len(za_dokaz), len(preneseni)), flush=True)
    snimi()

    if not za_dokaz:
        # Nista nije mereno -- ni baza. To mora da stoji U LINIJI VERDIKTA, ne
        # u fusnoti, jer bi se inace citalo kao svez dokaz.
        rezultat.update(verdikt="DOKAZANO", u_toku=False,
                        zavrsen=datetime.datetime.now().isoformat(timespec="seconds"))
        snimi()
        print("=== DOKAZANO (sve preneseno iz knjige -- nista nije mereno "
              "sada, ni baza) ===")
        return 0

    # --- KAPIJA: baza mora biti zelena pre prve mutacije --------------------
    potrebne = sorted({_suite_za(t) for _, _, t, _, _ in za_dokaz} | {SUITE_ALL})
    for suite in potrebne:
        ok, opis = _baza_zelena(suite)
        print("BAZNO: %s" % opis, flush=True)
        rezultat.setdefault("bazno", []).append(opis)
        if not ok:
            rezultat.update(verdikt="NIJE DOKAZANO", u_toku=False,
                            zavrsen=datetime.datetime.now().isoformat(timespec="seconds"))
            rezultat["problemi"].append(("(baza)", opis))
            snimi()
            print("STOP: baza nije zelena. Dokaz bi merio crveno koje sabotaza "
                  "nije izazvala.", file=sys.stderr)
            return 2

    pre = _otisak()
    print("potpis izvora: %s" % pre, flush=True)
    rezultat["potpis_pre"] = pre

    crvenih, lose, poznati = 0, [], []
    dokazani_sada = {}
    for ime, fajl, ocekTest, ocekTvrdnja, _unos in za_dokaz:
        p = _pusti(sys.executable, "tools/sabotaza.py", ime)
        if p.returncode != 0:
            lose.append((ime, "APPLY-FAIL -- v. sabotaza.py --proveri-sidra"))
            print("%-46s APPLY-FAIL" % ime, flush=True)
            rezultat["obradjeno"] += 1
            rezultat["stavke"].append({"ime": ime, "stanje": "APPLY-FAIL",
                                       "ok": False, "mereno": True})
            snimi()
            continue

        suite = _suite_za(ocekTest)
        pali, greska = [], ""
        try:
            run = _pusti(sys.executable, "tools/run_vba.py", "--suite", suite)
            pali = _pali(run.stdout + run.stderr, suite)
        except subprocess.TimeoutExpired:
            greska = "TIMEOUT suite"
        except KeyboardInterrupt:
            greska = "PREKID"
        finally:
            # Izvor se vraca i kad je run pukao -- inace radni tree ostaje
            # namerno pokvaren.
            v = _pusti(sys.executable, "tools/sabotaza.py", "--vrati")

        sada = _otisak()
        rezultat["obradjeno"] += 1
        if v.returncode != 0 or sada != pre:
            print("%-46s REVERT-FAIL (potpis %s)" % (ime, sada), flush=True)
            print("STOP: izvor nije vracen u pocetno stanje. Sve mereno posle "
                  "ovoga islo bi nad pokvarenim kodom.", file=sys.stderr)
            lose.append((ime, "REVERT-FAIL"))
            rezultat["stavke"].append({"ime": ime, "stanje": "REVERT-FAIL",
                                       "ok": False, "mereno": True})
            snimi()
            break

        if greska:
            lose.append((ime, greska))
            print("%-46s %s" % (ime, greska), flush=True)
            rezultat["stavke"].append({"ime": ime, "stanje": greska,
                                       "ok": False, "mereno": True})
            snimi()
            if greska == "PREKID":
                break
            continue

        if not pali:
            lose.append((ime, "NE OBARA NISTA"))
            stanje = "NE OBARA NISTA"
        else:
            crvenih += 1
            kljuc = _kljuc_testa(ocekTest, suite)
            imena = sorted({p0 for p0, _ in pali})
            # Poredi se SAMO ono sto je palo u njenom testu. Siroka
            # sabotaza obori i druge testove, pa bi tvrdnja iz TUDJEG
            # testa inace mogla da je "potvrdi".
            poruke = " | ".join(p1 for p0, p1 in pali if p0 == kljuc)
            if kljuc not in imena:
                stanje = "NE OBARA SVOJ TEST, nego: " + ", ".join(imena)
                lose.append((ime, stanje))
            elif not ocekTvrdnja:
                stanje = "KATALOG NEMA TVRDNJU -- nema sta da se poredi"
                lose.append((ime, stanje))
            elif ocekTvrdnja.lower() not in poruke.lower():
                # Pravi test a pogresna tvrdnja NIJE dokaz: ciljana tvrdnja
                # mozda nije ni izvrsena (AssertEq puca na prvom padu).
                stanje = ("PALA DRUGA TVRDNJA: " + poruke[:120])
                lose.append((ime, stanje))
            elif len(imena) > 1:
                stanje = "OK (uz jos %d testa)" % (len(imena) - 1)
                dokazani_sada[ime] = kljucevi[ime]
            else:
                stanje = "OK"
                dokazani_sada[ime] = kljucevi[ime]
        print("%-46s %s" % (ime, stanje), flush=True)
        rezultat["crvenih"] = crvenih
        rezultat["stavke"].append({"ime": ime, "stanje": stanje,
                                   "ok": ime in dokazani_sada, "mereno": True})
        snimi()

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
    izabrana_imena = {s[0] for s in stavke}
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
    print("\ncrvenih: %d / sabotaza: %d%s%s"
          % (crvenih, len(stavke),
             " (preneseno: %d)" % len(preneseni) if preneseni else "",
             " (priznatih: %d)" % len(pokriveni) if pokriveni else ""))
    print("izvor pre/posle: %s / %s -> %s"
          % (pre, posle, "IDENTICAN" if pre == posle else "RAZLIKA!"))
    for ime, sta in poznati:
        print(" POZNATO: %s -> %s" % (ime, sta))
    for ime, sta in lose:
        print(" PROBLEM: %s -> %s" % (ime, sta))

    # Preneseni ulaze u IMENILAC i u BROJILAC: oni jesu dokazani, samo ne sada.
    # Bez toga bi svaki run sa knjigom zavrsio kao NIJE DOKAZANO.
    ok = (not lose and pre == posle
          and crvenih + len(preneseni) >= len(stavke) - len(pokriveni))

    if ok and dokazani_sada:
        knjiga.update({ime: {"kljuc": k, "verdikt": "OK", "commit": commit,
                             "vreme": datetime.datetime.now().isoformat(
                                 timespec="seconds")}
                       for ime, k in dokazani_sada.items()})
        _upisi_knjigu(knjiga)

    rezultat.update(verdikt="DOKAZANO" if ok else "NIJE DOKAZANO",
                    u_toku=False, potpis_posle=posle,
                    crvenih=crvenih, poznatih=len(pokriveni),
                    problemi=[list(x) for x in lose],
                    poznati=[list(x) for x in poznati],
                    zavrsen=datetime.datetime.now().isoformat(timespec="seconds"))
    snimi()

    rep = ""
    if preneseni:
        rep = " (mereno %d, preneseno %d)" % (len(za_dokaz), len(preneseni))
    print("=== %s%s ===" % ("DOKAZANO" if ok else "NIJE DOKAZANO", rep))
    return 0 if ok else 1


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except KeyboardInterrupt:
        # Poslednja mreza: prekid izmedju dve mutacije ne sme da ostavi
        # pokvaren izvor.
        subprocess.run([sys.executable, "tools/sabotaza.py", "--vrati"], cwd=ROOT)
        print("\nprekinuto -- izvor vracen", file=sys.stderr)
        sys.exit(130)
