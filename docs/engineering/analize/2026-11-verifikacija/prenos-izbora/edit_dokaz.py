P = "tools/dokaz.py"
t = open(P).read()

def rep(old, new):
    global t
    n = t.count(old)
    assert n == 1, (n, old[:90])
    t = t.replace(old, new)

# --- docstring ---------------------------------------------------------------
rep('''    python tools/dokaz.py amb-pisac --grupe    # vise sabotaza u JEDNOM prolazu
    python tools/dokaz.py --grupe 8 --plan     # samo plan grupa, bez Excela
''', '''    python tools/dokaz.py amb-pisac --grupe    # vise sabotaza u JEDNOM prolazu
    python tools/dokaz.py --grupe 8 --plan     # samo plan grupa, bez Excela
    python tools/dokaz.py amb-pisac --izbor    # pusta samo SVOJ test, ne celu suite
    python tools/dokaz.py --izolacija          # koji ciljani testovi ne prolaze SAMI
''')
rep('''Banka-suite ne ispisuje ime testa uz pad,''', '''5. IZBOR (`--izbor`) PUSTA SAMO CILJANE TESTOVE, ALI SA SVOJOM BAZOM. Prolaz
   suite-a (RunAllTests ~80 s, BFP ~150 s) zamenjuje prolaz samo testova
   clanova (`run_vba --test`). Tacka 1 se tada tvrdi nad ISTIM izborom: pre
   mutacije se isti spisak testova pusti nad cistim izvorom i mora biti zelen.
   Bez toga bi test koji SAM pada (zavisi od stanja koje mu ostavlja raniji
   test -- redosled u RunAllTests je nosec) "dokazao" svaku sabotazu nad sobom.
   Takav clan se ne gubi: grupa se rasturi na pojedinacna merenja, a test koji
   ni sam ne prolazi meri se CELOM suite-om, kao bez `--izbor`. Baza se pamti
   po izboru, pa vise sabotaza nad istim testom placa jednu.

Banka-suite ne ispisuje ime testa uz pad,''')

# --- izbor: koje suite ga primaju ------------------------------------------
rep('''def _otisak() -> str:''', '''def _suite_izbor() -> set:
    """Suite koje primaju `run_vba --test` -- iz SUITES (jedini izvor istine).

    Cita se iz kataloga, ne prepisuje: druga lista bi bila druga stvar koja
    moze da se razidje sa ulaznim tackama u VBA (v. vba_gate, nalaz IZBOR).
    """
    put = os.path.join(ROOT, "tools", "run_vba.py")
    spec = importlib.util.spec_from_file_location("_run_vba_za_dokaz", put)
    modul = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modul)
    return {ime for ime, meta in modul.SUITES.items() if meta.get("izbor")}


def _run_suite(suite: str, testovi=None) -> tuple:
    """Argumenti za `run_vba.py`: cela suite, ili samo `testovi` iz nje."""
    a = (sys.executable, "tools/run_vba.py", "--suite", suite)
    for test in testovi or ():
        a += ("--test", test)
    return a


def _otisak() -> str:''')

# --- baza: i nad izborom ----------------------------------------------------
rep('''def _baza_zelena(suite: str) -> tuple:
    """(je_zelena, opis). Tvrdi se, ne gleda se."""
    try:
        r = _pusti(sys.executable, "tools/run_vba.py", "--suite", suite)
    except subprocess.TimeoutExpired:
        return False, f"{suite}: timeout"
    izlaz = r.stdout + r.stderr
''', '''def _baza_zelena(suite: str, testovi=None) -> tuple:
    """(je_zelena, opis). Tvrdi se, ne gleda se.

    Uz `testovi` je baza samo taj izbor -- tacno ono sto ce prolaz sa mutacijom
    pustiti. Izbor sa nula provera nije zelen nego prazan: nista nije mereno.
    """
    try:
        r = _pusti(*_run_suite(suite, testovi))
    except subprocess.TimeoutExpired:
        return False, f"{suite}: timeout"
    izlaz = r.stdout + r.stderr
''')
rep('''    if int(m.group(2)) != 0:
        return False, f"{suite}: {m.group(0)}"
    return True, m.group(0)
''', '''    if int(m.group(2)) != 0:
        return False, f"{suite}: {m.group(0)}"
    if testovi and int(m.group(1)) == 0:
        return False, f"{suite}: izbor bez ijedne provere ({m.group(0)})"
    return True, m.group(0)
''')

# --- _prolaz ------------------------------------------------------------------
rep('''def _prolaz(clanovi: list, pre: str, grupno: bool) -> dict:''',
    '''def _prolaz(clanovi: list, pre: str, grupno: bool, izbor=frozenset(),
            baze=None) -> dict:''')
rep('''    Vraca {"stop": razlog ili "", "ishodi": [(stavka, priznato, stanje, crveno)],
           "ostatak": [(stavka, razlog)]}.
    """
    rez = {"stop": "", "ishodi": [], "ostatak": []}

    primenjeni = []''', '''    IZBOR (`izbor` = suite koje primaju `--test`): pusta se samo spisak testova
    clanova, a PRE prve mutacije isti spisak mora da bude zelen nad cistim
    izvorom (tacka 5 u docstring-u modula; `baze` pamti ishod po izboru). Ako
    nije, nista se ne primenjuje: grupa ide na pojedinacna merenja, a pojedinacan
    clan u "cela_suite" -- meri se celom suite-om, kao bez izbora.

    Vraca {"stop": razlog ili "", "ishodi": [(stavka, priznato, stanje, crveno)],
           "ostatak": [(stavka, razlog)], "cela_suite": [(stavka, razlog)]}.
    """
    rez = {"stop": "", "ishodi": [], "ostatak": [], "cela_suite": []}

    suite = clanovi[0]["suite"]
    testovi = None
    if suite in izbor:
        testovi = sorted({s["test"] for s in clanovi})
        kljuc = (suite, tuple(testovi))
        baze = {} if baze is None else baze
        if kljuc not in baze:
            baze[kljuc] = _baza_zelena(suite, testovi)
        ok, opis = baze[kljuc]
        if not ok:
            for s in clanovi:
                if grupno and len(clanovi) > 1:
                    rez["ostatak"].append((s, "baza izbora nije zelena -- %s" % opis))
                else:
                    rez["cela_suite"].append((s, opis))
            return rez

    primenjeni = []''')
rep('''    sam = len(primenjeni) == 1
    suite = primenjeni[0]["suite"]
    ocekivani = {s["kljuc"] for s in primenjeni}

    pali, greska = [], ""
    try:
        run = _pusti(sys.executable, "tools/run_vba.py", "--suite", suite)''', '''    sam = len(primenjeni) == 1
    ocekivani = {s["kljuc"] for s in primenjeni}

    pali, greska = [], ""
    try:
        # Izbor je isti kao u bazi (svi clanovi, i oni ciji se sidro nije
        # primenilo): baza je merila TAJ spisak, pa se meri on.
        run = _pusti(*_run_suite(suite, testovi))''')

# --- izolacija --------------------------------------------------------------
rep('''def _plan(stavke: list, cap: int) -> int:''', '''def _izolacija(stavke: list, izbor: set) -> list:
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
        ok, opis = _baza_zelena(s["suite"], [s["test"]])
        print("%-6s %-28s %-50s %s" % ("OK" if ok else "CRVEN", s["suite"],
                                       s["test"], "" if ok else opis), flush=True)
        if not ok:
            crveni.append((s["suite"], s["test"], opis))
    print("\\nizolacija: %d testova, %d ne prolazi samo" % (len(vidjeni), len(crveni)))
    return crveni


def _plan(stavke: list, cap: int) -> int:''')

# --- argumenti ---------------------------------------------------------------
rep('''    ap.add_argument("--plan", action="store_true",
                    help="ispisi podelu na grupe i stani (ne trazi Excel)")''', '''    ap.add_argument("--plan", action="store_true",
                    help="ispisi podelu na grupe i stani (ne trazi Excel)")
    ap.add_argument("--izbor", action="store_true",
                    help="pusti samo ciljane testove (run_vba --test) umesto cele "
                         "suite; baza se tvrdi nad istim izborom")
    ap.add_argument("--izolacija", action="store_true",
                    help="bez mutacija: pusti svaki ciljani test SAM i nabroji "
                         "one koji ne prolaze")''')

rep('''    if args.plan:
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
''', '''    if args.plan:
        return _plan(stavke, args.grupe)

    izbor = _suite_izbor() if (args.izbor or args.izolacija) else set()
    if args.izolacija:
        return 1 if _izolacija(stavke, izbor) else 0

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
        ok, opis = _baza_zelena(suite)
        cele_baze[suite] = (ok, opis)
        print("BAZNO: %s" % opis, flush=True)
        if not ok:
            print("STOP: baza nije zelena. Dokaz bi merio crveno koje sabotaza "
                  "nije izazvala.", file=sys.stderr)
            return 2
''')

rep('''    grupe = naprav_grupe(stavke, args.grupe)
    ponavljaju = []''', '''    grupe = naprav_grupe(stavke, args.grupe)
    ponavljaju = []
    baze_izbora = {}
    cela_suite = []''')

rep('''        ishod = _prolaz(g, pre, len(g) > 1)
        if len(g) > 1:''', '''        ishod = _prolaz(g, pre, len(g) > 1, izbor, baze_izbora)
        cela_suite.extend(ishod["cela_suite"])
        if len(g) > 1:''')

rep('''        for s in ponavljaju:
            ishod = _prolaz([s], pre, False)
            upisi(ishod["ishodi"])''', '''        for s in ponavljaju:
            ishod = _prolaz([s], pre, False, izbor, baze_izbora)
            cela_suite.extend(ishod["cela_suite"])
            upisi(ishod["ishodi"])''')

rep('''    if stop:
        if stop.startswith("REVERT-FAIL"):''', '''    # Test koji SAM ne prolazi meri se celom suite-om -- uz bazu CELE suite,
    # koja se tvrdi jednom po suite-u i tek kad zatreba.
    if cela_suite and not stop:
        print("MERI SE CELOM SUITE-OM: %d (ciljani test sam ne prolazi -- v. "
              "--izolacija)" % len(cela_suite), flush=True)
        for s, razlog in cela_suite:
            if s["suite"] not in cele_baze:
                cele_baze[s["suite"]] = _baza_zelena(s["suite"])
                print("BAZNO: %s" % cele_baze[s["suite"]][1], flush=True)
            ok, opis = cele_baze[s["suite"]]
            if not ok:
                lose.append((s["ime"], "NIJE IZMERENA: ni izbor ni cela suite "
                                       "nisu zeleni -- %s" % opis))
                print("%-46s NIJE IZMERENA: baza cele suite nije zelena"
                      % s["ime"], flush=True)
                continue
            ishod = _prolaz([s], pre, False)
            upisi(ishod["ishodi"])
            if ishod["stop"]:
                stop = ishod["stop"]
                break

    if stop:
        if stop.startswith("REVERT-FAIL"):''')

rep('''    if grupno_priznatih:
        print("grupno izmereno: %d / %d''', '''    if izbor:
        print("izbor: baza i mutacija nad istim spiskom testova; celom suite-om "
              "mereno %d (test sam ne prolazi)" % len(cela_suite))
    if grupno_priznatih:
        print("grupno izmereno: %d / %d''')

open(P, "w").write(t)
print("ok")
