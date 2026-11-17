"""Diferencijalni dokaz za P1: stara i nova _dodela_tudjoj_proceduri daju ISTO
nad (1) celim pravim katalogom i (2) fuzz zamenama koje namerno pogadjaju pravilo."""
import importlib.util, os, random, sys, time
S = os.path.dirname(os.path.abspath(__file__))
def load(tag):
    spec = importlib.util.spec_from_file_location("sab_" + tag, os.path.join(S, tag, "tools", "sabotaza.py"))
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m); return m
B, P = load("base"), load("proto")
assert B.SABOTAZE == P.SABOTAZE
src = B.SRC_VBA
kes = {}
def tekst(f):
    if f not in kes: kes[f] = B._procitaj(os.path.join(src, f))[0]
    return kes[f]

# (1) pravi katalog
razl = 0; ne_none = 0
tb = tp = 0.0
for ime, (f, staro, novo, _t, _tv) in B.SABOTAZE.items():
    t = tekst(f)
    t0 = time.perf_counter(); a = B._dodela_tudjoj_proceduri(t, staro, novo); tb += time.perf_counter() - t0
    t0 = time.perf_counter(); b = P._dodela_tudjoj_proceduri(t, staro, novo); tp += time.perf_counter() - t0
    if a != b: razl += 1; print("RAZLIKA", ime, a, b)
    if a is not None: ne_none += 1
print(f"(1) katalog: {len(B.SABOTAZE)} unosa, razlika={razl}, ne-None={ne_none}; staro {tb:.2f}s -> novo {tp:.3f}s")

# (2) fuzz
rnd = random.Random(20261115)
fajlovi = sorted({v[0] for v in B.SABOTAZE.values()})
N = int(sys.argv[1]) if len(sys.argv) > 1 else 6000
razl = 0; ne_none = 0; izvrseno = 0; tb = tp = 0.0
vrste = {}
for it in range(N):
    f = rnd.choice(fajlovi); t = tekst(f)
    redovi = t.split("\n")
    procs = B._procedure_po_redu(t)
    if not procs: continue
    i = rnd.randrange(1, len(redovi) - 6)
    k = rnd.randint(1, 5)
    blok = redovi[i:i + k]
    staro = "\n".join(blok)
    oblik = rnd.choice(["sa_nl", "bez_nl", "pola_reda"])
    if oblik == "sa_nl": staro += "\n"
    elif oblik == "pola_reda" and len(blok[-1]) > 4: staro = staro[:len(staro) - rnd.randint(1, len(blok[-1]) - 1)]
    if t.count("\n" + staro) < 1: continue
    pime, pv, _a, _b = rnd.choice(procs)
    ime = rnd.choice([pime, pime.lower(), pime.upper()])
    dodela = rnd.choice(["    %s = 1", "%s = True", "    Set %s = Nothing", "    %s == 2", "    x = %s", "    %s.Foo = 1"]) % ime
    delovi = []
    vrsta = rnd.choice(["samo_dodela", "dim_pa_dodela", "static_pa_dodela", "const_pa_dodela", "zatvori_otvori",
                        "otvori_function", "zadrzi_blok_pa_dodela", "end_pa_dodela", "dim_prelomljen",
                        "property_otvori", "bez_dodele"])
    if vrsta == "samo_dodela": delovi = [dodela]
    elif vrsta == "dim_pa_dodela": delovi = ["    Dim %s As Long" % ime, dodela]
    elif vrsta == "static_pa_dodela": delovi = ["    Static a As Long, %s As Boolean" % ime, dodela]
    elif vrsta == "const_pa_dodela": delovi = ["    Const %s As Long = 1" % ime, dodela]
    elif vrsta == "zatvori_otvori": delovi = ["End Sub", "Private Sub Fuzz_%d(ByVal %s As Long)" % (it, rnd.choice([ime, "q"])), dodela]
    elif vrsta == "otvori_function": delovi = ["End Function", "Function %s() As Long" % rnd.choice([ime, "FuzzF"]), dodela]
    elif vrsta == "zadrzi_blok_pa_dodela": delovi = blok + [dodela]
    elif vrsta == "end_pa_dodela": delovi = [rnd.choice(["End Sub", "End Function", "End Property"]), dodela]
    elif vrsta == "dim_prelomljen": delovi = ["    Dim a As Long, _", "        %s As Long" % ime, dodela]
    elif vrsta == "property_otvori": delovi = ["End Sub", "Property Let %s(v As Long)" % rnd.choice([ime, "FuzzP"]), dodela]
    else: delovi = ["    ' nista", "    Debug.Print 1"]
    novo = "\n".join(delovi) + rnd.choice(["\n", "", "\n\n"])
    t0 = time.perf_counter(); a = B._dodela_tudjoj_proceduri(t, staro, novo); tb += time.perf_counter() - t0
    t0 = time.perf_counter(); b = P._dodela_tudjoj_proceduri(t, staro, novo); tp += time.perf_counter() - t0
    izvrseno += 1
    v = vrste.setdefault(vrsta, [0, 0]); v[0] += 1
    if a is not None: ne_none += 1; v[1] += 1
    if a != b:
        razl += 1
        if razl <= 5: print("RAZLIKA", f, vrsta, oblik, repr(staro[:60]), repr(novo[:80]), a, b)
print(f"(2) fuzz: {izvrseno} zamena, razlika={razl}, ne-None={ne_none} ({100*ne_none/max(1,izvrseno):.0f}%); staro {tb:.2f}s -> novo {tp:.3f}s")
for k, (n, nn) in sorted(vrste.items()): print(f"    {k:24s} {n:5d} zamena, nalaz u {nn}")
sys.exit(1 if razl else 0)
