import sys, os, importlib.util, collections
sys.dont_write_bytecode = True
ROOT = sys.argv[1]
spec = importlib.util.spec_from_file_location("_dok", os.path.join(ROOT, "tools", "dokaz.py")); dok = importlib.util.module_from_spec(spec); spec.loader.exec_module(dok)
sab = dok._modul_sabotaza()
kes = {}; orig = sab._suita_testa
def memo(t):
    if t not in kes: kes[t] = orig(t)
    return kes[t]
sab._suita_testa = memo
stavke = dok._izaberi(sab, [])
def blok_plan(ss, cap):
    # testovi po opadajucoj visestrukosti; blok = cap testova; sloj = po jedna sabotaza svakog testa, bez sudara procedure/kljuca
    po_testu = collections.defaultdict(list)
    for s in ss: po_testu[s["test"]].append(s)
    testovi = sorted(po_testu, key=lambda t: -len(po_testu[t]))
    blokovi = [testovi[i:i+cap] for i in range(0, len(testovi), cap)]
    prolaza = 0
    for b in blokovi:
        preostalo = {t: list(po_testu[t]) for t in b}
        while any(preostalo.values()):
            sloj = []
            for t in b:
                for i, s in enumerate(preostalo[t]):
                    if all(x["proc"] != s["proc"] and x["kljuc"] != s["kljuc"] for x in sloj):
                        sloj.append(s); preostalo[t].pop(i); break
            assert sloj
            prolaza += 1
    return len(blokovi), prolaza
for cap in (4, 6, 8, 12):
    tot_b = tot_p = 0; det = []
    for su in (dok.SUITE_ALL, dok.SUITE_BFP):
        ss = [s for s in stavke if s["suite"] == su]
        b, p = blok_plan(ss, cap); tot_b += b; tot_p += p; det.append("%s: baza %d, prolaza %d" % (su.replace("Run",""), b, p))
    print("cap %2d blok-plan: baza %d + prolaza %d = %d poziva | %s" % (cap, tot_b, tot_p, tot_b + tot_p, " ; ".join(det)))
