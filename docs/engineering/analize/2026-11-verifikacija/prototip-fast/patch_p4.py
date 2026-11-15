"""P4 (prototip): kes cinjenica PO FAJLU, kljuc = (ime, sha1 sadrzaja) + otisak alata.
Globalne provere postaju O(izmenjeni fajlovi) + spajanje. Radi nad KOPIJOM."""
import sys, os
tools = sys.argv[1]
def patch(ime, zamene):
    p = os.path.join(tools, ime)
    s = open(p, encoding="utf-8").read()
    for old, new in zamene:
        assert s.count(old) == 1, (ime, s.count(old), old[:80])
        s = s.replace(old, new)
    open(p, "w", encoding="utf-8").write(s)

# --- vba_gate.skeniraj: telo petlje u funkciju po fajlu -----------------------
patch("vba_gate.py", [('''    for fajl in sorted(os.listdir(src_dir)):
        if not fajl.endswith(TEKST_NASTAVCI):
            continue
        with io.open(os.path.join(src_dir, fajl), encoding="ascii",
                     errors="replace", newline="") as fh:
            tekst = fh.read().replace("\\r\\n", "\\n")
        dubina = 0
        for red, _broj in izjave(tekst):
            if _USLOV_POC.match(red):
                dubina += 1
                continue
            if _USLOV_KRAJ.match(red):
                dubina = max(0, dubina - 1)
                continue
            d = razlagac(red)
            if not d:
                continue
            d = dict(d, fajl=fajl, bas=fajl.endswith(".bas"),
                     uslovna=dubina > 0)
            deklaracije.setdefault(d["ime"], []).append(d)
            if je_ulazna_tacka(d) and IME_SUITE.match(d["ime"]):
                po_konvenciji.setdefault(d["ime"], fajl)
    return deklaracije, po_konvenciji
''', '''    for fajl in sorted(os.listdir(src_dir)):
        if not fajl.endswith(TEKST_NASTAVCI):
            continue
        spoji_skenirano(deklaracije, po_konvenciji,
                        skeniraj_fajl(src_dir, fajl, razlagac, izjave))
    return deklaracije, po_konvenciji


def skeniraj_fajl(src_dir, fajl, razlagac, izjave) -> list:
    """Razlozene deklaracije JEDNOG fajla -- isto telo petlje kao ranije."""
    with io.open(os.path.join(src_dir, fajl), encoding="ascii",
                 errors="replace", newline="") as fh:
        tekst = fh.read().replace("\\r\\n", "\\n")
    out, dubina = [], 0
    for red, _broj in izjave(tekst):
        if _USLOV_POC.match(red):
            dubina += 1
            continue
        if _USLOV_KRAJ.match(red):
            dubina = max(0, dubina - 1)
            continue
        d = razlagac(red)
        if not d:
            continue
        out.append(dict(d, fajl=fajl, bas=fajl.endswith(".bas"),
                        uslovna=dubina > 0))
    return out


def spoji_skenirano(deklaracije, po_konvenciji, lista) -> None:
    for d in lista:
        deklaracije.setdefault(d["ime"], []).append(d)
        if je_ulazna_tacka(d) and IME_SUITE.match(d["ime"]):
            po_konvenciji.setdefault(d["ime"], d["fajl"])
''')])

# --- vba_check ----------------------------------------------------------------
patch("vba_check.py", [
# collect_arities -> sirovo "seen" po fajlovima + filter
('''    seen: dict[str, set[tuple[int, float]]] = defaultdict(set)
    for path in files:
        with open(path, "r", encoding="ascii", errors="replace") as fh:
            lines = fh.read().replace("\\r\\n", "\\n").split("\\n")
        i = 0
''', '''    return _arnost_jednoznacna(_arnost_vidjeno(files))


def _arnost_jednoznacna(seen: dict) -> dict:
    return {name: next(iter(v)) for name, v in seen.items() if len(v) == 1}


def _arnost_vidjeno(files: list[str]) -> dict:
    seen: dict[str, set[tuple[int, float]]] = defaultdict(set)
    for path in files:
        with open(path, "r", encoding="ascii", errors="replace") as fh:
            lines = fh.read().replace("\\r\\n", "\\n").split("\\n")
        i = 0
'''),
('''            i += 1
    return {name: next(iter(v)) for name, v in seen.items() if len(v) == 1}
''', '''            i += 1
    return seen
'''),
# amb_bind: po fajlu
('''        with open(os.path.join(src_dir, ime_f), "r", encoding="ascii",
                  errors="replace") as fh:
            tekst = fh.read()
        proc = None
        for izjava, nred in logicke_izjave(tekst):
            go = izjava.strip()
            if _AMB_END.match(go):
                proc = None
                continue
            d = deklaracija_procedure(go)
            if d:
                proc = d["ime"]
                continue
            # String literali se skidaju: poruka o gresci sme da pomene ime.
            if not _BIND_POZIV.search(_BIND_STRING.sub('""', go)):
                continue
            if (modul, proc) in AMB_BIND_DOZVOLJENI:
                continue
            out.append(Finding(
''', '''        for nred, proc in _amb_pozivi_fajla(os.path.join(src_dir, ime_f)):
            if (modul, proc) in AMB_BIND_DOZVOLJENI:
                continue
            out.append(Finding(
'''),
('''def check_amb_bind_vlasnik(src_dir: str | None = None) -> list[Finding]:''',
'''def _amb_pozivi_fajla(put: str) -> list:
    """[(red, procedura)] svih poziva BindSourceDocument u fajlu -- kesira se."""
    k = _KES_AMB.get(put)
    if k is not None:
        return k
    with open(put, "r", encoding="ascii", errors="replace") as fh:
        tekst = fh.read()
    out, proc = [], None
    for izjava, nred in logicke_izjave(tekst):
        go = izjava.strip()
        if _AMB_END.match(go):
            proc = None
            continue
        d = deklaracija_procedure(go)
        if d:
            proc = d["ime"]
            continue
        # String literali se skidaju: poruka o gresci sme da pomene ime.
        if not _BIND_POZIV.search(_BIND_STRING.sub('""', go)):
            continue
        out.append((nred, proc))
    return out


_KES_AMB: dict = {}


def check_amb_bind_vlasnik(src_dir: str | None = None) -> list[Finding]:'''),
# main: kes
('''    rc_kat = check_katalog_sabotaza(args.hook)
    rc_grupe = check_dokaz_grupe(args.hook)
    rc_popis = check_popis_suita(args.hook)
''', '''    kes = _kes_ucitaj()
    svi = vba_files([])
    cinj = {}
    for p in svi:
        with open(p, "rb") as fh:
            kljuc = (os.path.basename(p), hashlib.sha1(fh.read()).hexdigest())
        c = kes["fajl"].get(kljuc)
        if c is None:
            c = _cinjenice_fajla(p)
            kes["prljav"] = True
        cinj[p] = (kljuc, c)
        if "amb" in c:
            _KES_AMB[p] = c["amb"]
    rc_kat = _katalog_kroz_kes(args.hook, kes, cinj)
    rc_grupe = check_dokaz_grupe(args.hook)
    rc_popis = check_popis_suita(args.hook, cinj)
'''),
('''    defined = collect_definitions(vba_files([]))
    moduli = collect_module_members(vba_files([]))
    arities = collect_arities(vba_files([]))
''', '''    defined, moduli, vidjeno = set(), {}, defaultdict(set)
    for p in svi:
        c = cinj[p][1]
        defined |= c["defs"]
        moduli.update(c["clanovi"])
        for ime, s in c["arnost"].items():
            vidjeno[ime] |= s
    arities = _arnost_jednoznacna(vidjeno)
    _kes_snimi(kes, cinj)
'''),
('''def check_popis_suita(tiho: bool = False) -> int:''', '''def check_popis_suita(tiho: bool = False, cinj: dict = None) -> int:'''),
('''        nalazi = modul.popis_problemi(
            razlagac=(deklaracija_procedure, logicke_izjave))
''', '''        skenirano = None
        if cinj is not None:
            dekl, konv = {}, {}
            for p in sorted(cinj):
                modul.spoji_skenirano(dekl, konv, cinj[p][1]["sken"])
            skenirano = (dekl, konv)
        nalazi = modul.popis_problemi(
            skenirano=skenirano,
            razlagac=(deklaracija_procedure, logicke_izjave))
'''),
('''def main(argv: list[str]) -> int:''', '''# --- PROTOTIP P4: kes cinjenica po fajlu -------------------------------------
import hashlib
import pickle

_ALATI = ("vba_check.py", "sabotaza.py", "vba_gate.py", "run_vba.py", "dokaz.py")


def _otisak_alata() -> str:
    h = hashlib.sha1()
    ovde = os.path.dirname(os.path.abspath(__file__))
    for ime in _ALATI:
        p = os.path.join(ovde, ime)
        if os.path.exists(p):
            with open(p, "rb") as fh:
                h.update(fh.read())
        h.update(b"|")
    return h.hexdigest()


def _kes_put() -> str:
    return os.path.join(tempfile.gettempdir(), "vba_check_kes_%s.pkl"
                        % hashlib.sha1(ROOT.encode()).hexdigest()[:12])


def _kes_ucitaj() -> dict:
    alat = _otisak_alata()
    try:
        with open(_kes_put(), "rb") as fh:
            k = pickle.load(fh)
        if k.get("alat") == alat:
            k["prljav"] = False
            return k
    except Exception:
        pass
    return {"alat": alat, "fajl": {}, "katalog": {}, "prljav": True}


def _kes_snimi(kes: dict, cinj: dict) -> None:
    if not kes.get("prljav"):
        return
    kes["fajl"] = {k: c for k, c in cinj.values()}     # samo zivi fajlovi
    kes["prljav"] = False
    tmp = _kes_put() + ".%d.tmp" % os.getpid()
    with open(tmp, "wb") as fh:
        pickle.dump(kes, fh, protocol=pickle.HIGHEST_PROTOCOL)
    os.replace(tmp, _kes_put())


def _cinjenice_fajla(p: str) -> dict:
    """Sve sto globalne provere citaju iz JEDNOG fajla -- istim funkcijama kao pre."""
    gate = _gate_modul()
    return {
        "defs": collect_definitions([p]),
        "clanovi": collect_module_members([p]),
        "arnost": dict(_arnost_vidjeno([p])),
        "sken": gate.skeniraj_fajl(os.path.dirname(p), os.path.basename(p),
                                   deklaracija_procedure, logicke_izjave),
        "amb": _amb_pozivi_fajla(p),
    }


_gate_kes = []


def _gate_modul():
    if not _gate_kes:
        put = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vba_gate.py")
        spec = importlib.util.spec_from_file_location("_gate_za_cinj", put)
        m = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(m)
        _gate_kes.append(m)
    return _gate_kes[0]


def _katalog_kroz_kes(tiho: bool, kes: dict, cinj: dict) -> int:
    """Presuda kataloga zavisi SAMO od sabotaza.py i fajlova koje on cita."""
    otisci = {os.path.basename(p): k[1] for p, (k, _c) in cinj.items()}
    stari = kes["katalog"].get(tiho)
    if stari and all(otisci.get(f) == o for f, o in stari["ulaz"].items()):
        sys.stdout.write(stari["out"])
        sys.stderr.write(stari["err"])
        return stari["rc"]
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        rc = check_katalog_sabotaza(tiho)
    sys.stdout.write(out.getvalue())
    sys.stderr.write(err.getvalue())
    put = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sabotaza.py")
    if os.path.exists(put) and "ne ucitava" not in err.getvalue():
        spec = importlib.util.spec_from_file_location("_sab_za_kes", put)
        m = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(m)
        fajlovi = {v[0] for v in m.SABOTAZE.values()} | set(m._TEST_FAJLOVI)
        kes["katalog"][tiho] = {"ulaz": {f: otisci.get(f) for f in fajlovi},
                                "out": out.getvalue(), "err": err.getvalue(),
                                "rc": rc}
        kes["prljav"] = True
    return rc


def main(argv: list[str]) -> int:'''),
])
print("patched", tools)
