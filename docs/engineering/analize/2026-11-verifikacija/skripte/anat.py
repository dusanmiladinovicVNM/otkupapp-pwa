import sys, os, importlib.util, collections, re, json
sys.dont_write_bytecode = True
ROOT = sys.argv[1]
def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m
dok = load("_dok", os.path.join(ROOT, "tools", "dokaz.py"))
sab = dok._modul_sabotaza()
stavke = dok._izaberi(sab, [])
N = len(stavke)
print("ukupno sabotaza:", N)
C = collections.Counter
# suite
by_suite = C(s["suite"] for s in stavke)
print("po suite:", dict(by_suite))
# fajl
by_fajl = C(s["fajl"] for s in stavke)
print("razlicitih fajlova:", len(by_fajl))
print("top 20 fajlova:")
for f, n in by_fajl.most_common(20):
    print("   %4d  %s" % (n, f))
cum = 0
for i,(f,n) in enumerate(by_fajl.most_common(),1):
    cum += n
    if i in (1,3,5,10,20): print("   top %d fajlova = %d (%.0f%%)" % (i, cum, 100*cum/N))
# procedura
by_proc = C(s["proc"] for s in stavke)
print("razlicitih procedura:", len(by_proc), "; None:", by_proc.get(None,0))
print("deklaracije:", sum(n for p,n in by_proc.items() if p and p.endswith("<deklaracije>")))
print("top 20 procedura:")
for p, n in by_proc.most_common(20):
    print("   %4d  %s" % (n, p))
hist = C(by_proc.values())
print("histogram (sabotaza po proceduri -> broj procedura):", sorted(hist.items()))
print("procedura sa 1 sabotazom:", hist[1], "; sa >=5:", sum(v for k,v in hist.items() if k>=5), "; sa >=10:", sum(v for k,v in hist.items() if k>=10))
# test
by_test = C((s["suite"], s["test"]) for s in stavke)
print("razlicitih testova:", len(by_test))
for su in by_suite:
    tt = {t:n for (s_,t),n in by_test.items() if s_==su}
    h = C(tt.values())
    print("  ", su, "testova", len(tt), "sa 1 sab:", h[1], "sa 2:", h[2], "3-5:", sum(v for k,v in h.items() if 3<=k<=5), "6-9:", sum(v for k,v in h.items() if 6<=k<=9), ">=10:", sum(v for k,v in h.items() if k>=10), "max:", max(tt.values()))
    sab_in = lambda lo,hi: sum(k*v for k,v in h.items() if lo<=k<=hi)
    print("      sabotaza u testovima sa 1:", sab_in(1,1), "2:", sab_in(2,2), "3-5:", sab_in(3,5), "6-9:", sab_in(6,9), ">=10:", sab_in(10,999))
print("top 25 testova:")
for (su,t), n in by_test.most_common(25):
    print("   %4d  %-26s %s" % (n, su.replace("Run",""), t))
# (test, tvrdnja)
by_tt = C((s["suite"], s["test"], s["tvrdnja"].strip().lower()) for s in stavke)
dup = {k:n for k,n in by_tt.items() if n>1}
print("razlicitih (test,tvrdnja):", len(by_tt), "; kljuceva sa >1 sabotazom:", len(dup), "; sabotaza u njima:", sum(dup.values()), "; visak:", sum(n-1 for n in dup.values()))
print("histogram (sabotaza po (test,tvrdnja)):", sorted(C(by_tt.values()).items()))
for k,n in sorted(dup.items(), key=lambda x:-x[1])[:15]:
    print("   %3d  %s | %s" % (n, k[1], k[2][:80]))
# po suite dup
for su in by_suite:
    d = {k:n for k,n in dup.items() if k[0]==su}
    print("   ", su, "dup kljuceva", len(d), "sabotaza", sum(d.values()))
# kljuc (dokaz key)
by_kljuc = C((s["suite"], s["kljuc"]) for s in stavke)
print("razlicitih dokaz-kljuceva:", len(by_kljuc))
# sidro: isti (fajl, sidro)
by_sidro = C((f, sidro) for ime,(f,sidro,novo,t,tv) in sab.SABOTAZE.items())
ds = {k:n for k,n in by_sidro.items() if n>1}
print("razlicitih (fajl,sidro):", len(by_sidro), "; sidara sa >1 sabotazom:", len(ds), "; sabotaza u njima:", sum(ds.values()))
# isti prvi red sidra (pocetni red)
# compute start line of each anchor and line span
spans = {}
for ime,(f,sidro,novo,t,tv) in sab.SABOTAZE.items():
    dok._procedura(f, sidro, sab)
    kes = dok._PROC_KES.get(f)
    if not kes: continue
    tekst, _ = kes
    poz = tekst.find("\n"+sidro)
    if poz < 0: continue
    a = tekst[:poz+1].count("\n")
    b = a + sidro.rstrip("\n").count("\n")
    spans[ime] = (f, a, b)
by_start = C((f,a) for f,a,b in spans.values())
print("sidra nadjena:", len(spans), "; razlicitih (fajl,prvi red):", len(by_start), "; pocetnih redova sa >1:", sum(1 for n in by_start.values() if n>1), "sabotaza:", sum(n for n in by_start.values() if n>1))
# overlap of line ranges
byf = collections.defaultdict(list)
for ime,(f,a,b) in spans.items(): byf[f].append((a,b,ime))
overl = set()
for f, L in byf.items():
    L.sort()
    for i in range(len(L)):
        for j in range(i+1, len(L)):
            if L[j][0] > L[i][1]: break
            overl.add(L[i][2]); overl.add(L[j][2])
print("sabotaza cije se sidro (opseg redova) preklapa sa nekim drugim:", len(overl))
# sidro/zamena velicina
ls = [sidro.count("\n") for ime,(f,sidro,novo,t,tv) in sab.SABOTAZE.items()]
lz = [novo.count("\n") for ime,(f,sidro,novo,t,tv) in sab.SABOTAZE.items()]
import statistics
print("sidro redova: median %s, mean %.1f, max %d; 1-redna: %d" % (statistics.median(ls), statistics.mean(ls), max(ls), sum(1 for x in ls if x<=1)))
print("zamena redova: median %s, mean %.1f, max %d" % (statistics.median(lz), statistics.mean(lz), max(lz)))
print("ukupno redova sidra+zamene u katalogu:", sum(ls)+sum(lz))
# poznati
print("POZNATI_NALAZI:", len(sab.POZNATI_NALAZI), dict(C(v if isinstance(v,str) else tuple(v) for v in sab.POZNATI_NALAZI.values())))
print("POZNATI_NALAZI_DOKAZ:", len(sab.POZNATI_NALAZI_DOKAZ), sab.POZNATI_NALAZI_DOKAZ)
# plans
for cap in (1,2,3,4,6,8,12,24,1000):
    g = dok.naprav_grupe(stavke, cap)
    bs = C(x[0]["suite"] for x in g)
    print("cap %4d: prolaza %4d  %s  max grupa %d" % (cap, len(g), dict(bs), max(len(x) for x in g)))
json.dump([{k:(v) for k,v in s.items()} for s in stavke], open(os.path.join(os.path.dirname(__file__), "stavke_%s.json" % os.path.basename(ROOT.rstrip("/"))), "w"))
