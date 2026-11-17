"""Diferencijalni dokaz P2: brze putanje daju ISTO sto i originalne petlje,
nad svakim fizickim i logickim redom src-vba + nad neprijateljskim redovima."""
import importlib.util, os, random, sys, glob
S = os.path.dirname(os.path.abspath(__file__))
def load(tag):
    sys.path.insert(0, os.path.join(S, tag, "tools"))
    spec = importlib.util.spec_from_file_location("vc_" + tag, os.path.join(S, tag, "tools", "vba_check.py"))
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    sys.path.pop(0); return m
B, P = load("base"), load(sys.argv[1] if len(sys.argv) > 1 else "proto")
files = B.vba_files([])
redovi = set()
for f in files:
    raw = open(f, "rb").read().decode("ascii", errors="replace").replace("\r\n", "\n").split("\n")
    redovi.update(raw)
    redovi.update(t for t, _ in B._logical_lines(raw))
rnd = random.Random(7)
abeceda = ['"', "'", ":", "(", ")", "[", "]", ",", " ", "_", "=", "a", "Sub ", "Function ", "x", "\t", ":=", '""', "' _", " _", "Property Get "]
fuzz = set()
for _ in range(200000):
    fuzz.add("".join(rnd.choice(abeceda) for _ in range(rnd.randint(0, 14))))
svi = list(redovi) + list(fuzz)
print(f"razlicitih redova iz src-vba: {len(redovi)}, fuzz redova: {len(fuzz)}")
for ime in ["_izjava_deo", "_strip_comment", "_split_statements", "_strip_strings", "_clean"]:
    fb, fp = getattr(B, ime), getattr(P, ime)
    razl = 0
    for r in svi:
        if fb(r) != fp(r):
            razl += 1
            if razl <= 3: print("  RAZLIKA", ime, repr(r), fb(r), fp(r))
    print(f"{ime:20s} poredjenja={len(svi)} razlika={razl}")
ab, apr = B.collect_arities(files), P.collect_arities([f for f in files])
print("collect_arities: imena", len(ab), "jednako:", ab == apr)
# logicke_izjave nad celim fajlovima
razl = 0
for f in files:
    t = open(f, encoding="ascii", errors="replace").read()
    if B.logicke_izjave(t) != P.logicke_izjave(t): razl += 1
print("logicke_izjave: fajlova", len(files), "razlika", razl)
