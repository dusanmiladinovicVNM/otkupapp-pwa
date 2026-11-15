import sys, os, importlib.util, collections
sys.dont_write_bytecode = True
ROOT = sys.argv[1]
spec = importlib.util.spec_from_file_location("_dok", os.path.join(ROOT, "tools", "dokaz.py")); dok = importlib.util.module_from_spec(spec); spec.loader.exec_module(dok)
sab = dok._modul_sabotaza()
# memoize _suita_testa to avoid 27 s
kes = {}
orig = sab._suita_testa
sab._suita_testa = lambda t: kes.setdefault(t, orig(t)) if t not in kes else kes[t]
stavke = dok._izaberi(sab, [])
IZB = {dok.SUITE_ALL, dok.SUITE_BFP}
for cap in (1, 2, 3, 4, 6, 8, 12, 24):
    g = dok.naprav_grupe(stavke, cap)
    baze = set(); prol = collections.Counter(); testova_u_prolazu = collections.Counter()
    for x in g:
        su = x[0]["suite"]; prol[su] += 1
        if su in IZB:
            k = (su, tuple(sorted({s["test"] for s in x})))
            baze.add(k)
            testova_u_prolazu[su] += len(k[1])
    bz = collections.Counter(k[0] for k in baze)
    tb = collections.Counter()
    for k in baze: tb[k[0]] += len(k[1])
    print("cap %2d: prolaza A=%d B=%d K=%d | baza izbora A=%d B=%d (ukupno %d) | run_vba poziva (bez ponavljanja) = %d | test-izvrsenja u prolazima A=%d B=%d, u bazama A=%d B=%d" % (
        cap, prol[dok.SUITE_ALL], prol[dok.SUITE_BFP], prol[dok.SUITE_BANKA], bz[dok.SUITE_ALL], bz[dok.SUITE_BFP], len(baze),
        len(g) + len(baze) + 1, testova_u_prolazu[dok.SUITE_ALL], testova_u_prolazu[dok.SUITE_BFP], tb[dok.SUITE_ALL], tb[dok.SUITE_BFP]))
# lower bound on passes: max sabotages per test per suite; per proc; per key
C = collections.Counter
for su in (dok.SUITE_ALL, dok.SUITE_BFP, dok.SUITE_BANKA):
    ss = [s for s in stavke if s["suite"] == su]
    print(su, "max po testu", max(C(s["test"] for s in ss).values()), "max po proceduri", max(C(s["proc"] for s in ss).values()), "max po kljucu", max(C(s["kljuc"] for s in ss).values()))
