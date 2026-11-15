"""Wall-clock unutar sabotaza.proveri_sidra (bez cProfile)."""
import importlib.util, sys, time, os, io, contextlib
from collections import defaultdict
tools = sys.argv[1]
t = time.perf_counter()
spec = importlib.util.spec_from_file_location("_sab", os.path.join(tools, "sabotaza.py"))
m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
t_imp = time.perf_counter() - t
T = defaultdict(float); N = defaultdict(int)
def wrap(name):
    f = getattr(m, name)
    def w(*a, **k):
        t = time.perf_counter()
        try: return f(*a, **k)
        finally:
            T[name] += time.perf_counter() - t; N[name] += 1
    setattr(m, name, w)
for n in ["_nalazi", "_imena_testova", "_tela_testova", "_dodela_tudjoj_proceduri",
          "_procedure_po_redu", "_lokalna_imena", "_pogodaka", "_tvrdnja_pripada",
          "_tvrdnja_ceo_literal", "_bfp_tela", "_procitaj"]:
    wrap(n)
t = time.perf_counter()
out = io.StringIO()
with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
    rc = m.proveri_sidra(True)
tot = time.perf_counter() - t
print(f"import sabotaza.py: {t_imp:.2f}s ; proveri_sidra: {tot:.2f}s rc={rc}; sabotaza={len(m.SABOTAZE)}")
for n in sorted(T, key=lambda k: -T[k]):
    print(f"{T[n]:7.2f}s x{N[n]:<6} {n}")
# struktura kataloga
from collections import Counter
c = Counter(v[0] for v in m.SABOTAZE.values())
src = m.SRC_VBA
sizes = {f: os.path.getsize(os.path.join(src, f)) for f in c if os.path.exists(os.path.join(src, f))}
print("razlicitih fajlova:", len(c))
tot_bytes = sum(c[f] * sizes.get(f, 0) for f in c)
print(f"suma (sabotaza x velicina fajla) = {tot_bytes/1e6:.1f} MB; suma velicina jedinstvenih = {sum(sizes.values())/1e6:.2f} MB")
for f, n in c.most_common(12):
    print(f"  {n:4d} x {f:38s} {sizes.get(f,0)/1024:7.0f} KB  -> {n*sizes.get(f,0)/1e6:6.1f} MB")
dod = sum(1 for v in m.SABOTAZE.values() if any(m._ZAMENA_DODELA.match(l) for l in v[2].split("\n")))
print("sabotaza cija ZAMENA ima bar jednu liniju-dodelu (_ZAMENA_DODELA):", dod, "od", len(m.SABOTAZE))
