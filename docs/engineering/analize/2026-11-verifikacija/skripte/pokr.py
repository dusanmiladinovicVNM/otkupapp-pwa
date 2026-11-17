import sys, os, importlib.util, collections, re
sys.dont_write_bytecode = True
ROOT = sys.argv[1]
def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m); return m
dok = load("_dok", os.path.join(ROOT, "tools", "dokaz.py"))
sab = dok._modul_sabotaza()
SRC = os.path.join(ROOT, "src-vba")
C = collections.Counter
# catalog claims per test
kat = collections.defaultdict(list)
for ime,(f,sidro,novo,t,tv) in sab.SABOTAZE.items():
    kat[t].append(tv.strip().lower())
def tela(fajl):
    tekst = sab._procitaj(os.path.join(SRC, fajl))[0]
    out = {}
    for m in re.finditer(r"^(?:Public |Private )?(?:Sub|Function) (\w+)", tekst, re.M):
        k = re.search(r"^End (?:Sub|Function)\b", tekst[m.start():], re.M)
        out[m.group(1)] = tekst[m.start(): m.start() + (k.end() if k else len(tekst))]
    return tekst, out
for fajl, pref in (("modTest.bas", "T_"), ("modBusinessFlowProTests.bas", "Test"), ("modTestBanka.bas", "T")):
    tekst, T = tela(fajl)
    tot_calls = 0; tot_distinct = 0; cov_calls = 0; cov_distinct = 0
    tests_with_assert = 0; tests_sab = 0; n_tests = 0
    dyn = 0
    helper_calls = 0
    per_test = []
    for name, telo in T.items():
        izrazi = sab._poruke_tvrdnji(telo)
        is_test = name.startswith(pref)
        if not is_test:
            helper_calls += len(izrazi); continue
        n_tests += 1
        if izrazi: tests_with_assert += 1
        msgs = []
        for iz in izrazi:
            pun, frag = sab._poruka_delovi(iz)
            if pun is None:
                dyn += 1
                msgs.append(("D", " ".join(frag).lower()))
            else:
                msgs.append(("S", pun.lower()))
        claims = kat.get(name, [])
        if claims: tests_sab += 1
        dist = set(msgs)
        tot_calls += len(msgs); tot_distinct += len(dist)
        def covered(m):
            kind, txt = m
            if fajl == "modBusinessFlowProTests.bas":
                return any(c == txt for c in claims) if kind == "S" else False
            # modTest: substring
            if kind == "S":
                return any(c in txt for c in claims)
            # dynamic: fragments in order within claim
            return any(all(fr in c for fr in txt.split("  ")) and txt and (txt[:12] in c) for c in claims)
        cc = sum(1 for m in msgs if covered(m)); cd = sum(1 for m in dist if covered(m))
        cov_calls += cc; cov_distinct += cd
        per_test.append((name, len(msgs), len(dist), cd, len(claims)))
    print("==", fajl)
    print("  procedura:", len(T), "; test-procedura (prefiks %s):" % pref, n_tests, "; sa bar 1 tvrdnjom:", tests_with_assert, "; sa bar 1 sabotazom:", tests_sab)
    print("  poziva tvrdnji u testovima:", tot_calls, "(od toga dinamicka poruka:", dyn, "); u helper-ima:", helper_calls)
    print("  razlicitih (test,poruka):", tot_distinct, "; pokrivenih sabotazom:", cov_distinct, "(%.1f%%)" % (100*cov_distinct/max(1,tot_distinct)))
    print("  poziva pokrivenih:", cov_calls, "(%.1f%%)" % (100*cov_calls/max(1,tot_calls)))
    print("  kataloskih tvrdnji za testove ovog modula:", sum(len(kat[n]) for n in T if n in kat))
    nosab = [p for p in per_test if p[4]==0 and p[1]>0]
    print("  testova sa tvrdnjama a BEZ sabotaze:", len(nosab), "; tvrdnji u njima:", sum(p[2] for p in nosab))
    # tests with sab: fraction of distinct covered
    withsab = [p for p in per_test if p[4]>0]
    if withsab:
        print("  u testovima SA sabotazom: razlicitih poruka", sum(p[2] for p in withsab), "pokrivenih", sum(p[3] for p in withsab))
        full = sum(1 for p in withsab if p[3] >= p[2])
        print("  testova gde je SVAKA razlicita poruka pokrivena:", full)
    # mismatch catalog vs cov (claims that matched nothing)
# RunOne count
mt = sab._procitaj(os.path.join(SRC, "modTest.bas"))[0]
print("RunOne poziva u modTest:", len(re.findall(r'^\s*RunOne\s+"', mt, re.M)), "; razlicitih:", len(set(re.findall(r'^\s*RunOne\s+"(\w+)"', mt, re.M))))
