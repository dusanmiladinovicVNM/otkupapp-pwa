"""Koliki deo kataloga bi PIT-ov inkrementalni kljuc ponistio po PR-u? (samo citanje, git objekti)"""
import subprocess, sys, os, re, hashlib, types, statistics
REPO = sys.argv[1]; N = int(sys.argv[2])
def git(*a):
    return subprocess.run(["git", "-C", REPO] + list(a), capture_output=True, check=True).stdout
_blob = {}
def blob(oid):
    if oid not in _blob:
        _blob[oid] = git("cat-file", "blob", oid).decode("latin-1").replace("\r\n", "\n")
    return _blob[oid]
def tree(c):
    out = {}
    for red in git("ls-tree", "-r", c, "src-vba", "tools/sabotaza.py").decode().splitlines():
        meta, put = red.split("\t", 1)
        out[put] = meta.split()[2]
    return out
_kat = {}
def katalog(oid):
    if oid not in _kat:
        m = types.ModuleType("sab_" + oid[:8]); m.__file__ = os.path.join(REPO, "tools", "sabotaza.py")
        exec(compile(blob(oid), "sabotaza.py", "exec"), m.__dict__)
        _kat[oid] = dict(m.SABOTAZE)
    return _kat[oid]
OTV = re.compile(r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property)\s+(?:(?:Get|Let|Set)\s+)?(\w+)", re.I)
ZAT = re.compile(r"^End\s+(?:Sub|Function|Property)\b", re.I)
_proc = {}
def procedure(oid):
    if oid not in _proc:
        redovi = blob(oid).split("\n"); out = []; ime = None; poc = 0
        for i, r in enumerate(redovi):
            t = r.strip(); m = OTV.match(t)
            if m: ime, poc = m.group(2), i
            elif ime and ZAT.match(t):
                out.append((ime, poc, i)); ime = None
        _proc[oid] = (redovi, out)
    return _proc[oid]
def h(s): return hashlib.sha1(s.encode("latin-1")).hexdigest()
def proc_sidra(oid, sidro):
    tekst = blob(oid); poz = tekst.find("\n" + sidro)
    if poz < 0: return None
    red = tekst[:poz + 1].count("\n"); redovi, pr = procedure(oid)
    for ime, a, b in pr:
        if a <= red <= b: return h("\n".join(redovi[a:b + 1]))
    return "DEKL:" + h("\n".join(redovi[:pr[0][1]] if pr else redovi))
TESTMOD = ["src-vba/modTest.bas", "src-vba/modTestBanka.bas", "src-vba/modBusinessFlowProTests.bas"]
def proc_testa(tr, test):
    for put in TESTMOD:
        if put not in tr: continue
        redovi, pr = procedure(tr[put])
        for ime, a, b in pr:
            if ime == test: return put, h("\n".join(redovi[a:b + 1]))
    return None, None
commits = git("log", "--first-parent", "-n", str(N), "--format=%H %cd", "--date=short").decode().splitlines()
print("%-9s %-10s %5s %5s %5s | %6s %6s %6s | %s" % ("commit", "datum", "kat", "nove", "izm", "K_fajl", "K_proc", "K_p+n", "dirnutih src-vba fajlova"))
rows = []
for red in commits:
    c, datum = red.split()
    p = git("rev-parse", c + "^1").decode().strip()
    tc, tp = tree(c), tree(p)
    if "tools/sabotaza.py" not in tc or "tools/sabotaza.py" not in tp: continue
    kc, kp = katalog(tc["tools/sabotaza.py"]), katalog(tp["tools/sabotaza.py"])
    dirnuti = sum(1 for f in tc if f.startswith("src-vba/") and tp.get(f) != tc[f])
    nove = izm = kf = kpr = 0
    for ime, tup in kc.items():
        if ime not in kp: nove += 1; continue
        if kp[ime] != tup: izm += 1; continue
        fajl, sidro, _n, test, _t = tup
        put = "src-vba/" + fajl
        tm_c, th_c = proc_testa(tc, test); tm_p, th_p = proc_testa(tp, test)
        isti_fajl = tc.get(put) == tp.get(put)
        isti_tm = tm_c and tm_c == tm_p and tc.get(tm_c) == tp.get(tm_p)
        if not (isti_fajl and isti_tm): kf += 1
        if isti_fajl and isti_tm: continue
        ps_c = proc_sidra(tc[put], sidro) if put in tc else None
        ps_p = proc_sidra(tp[put], sidro) if put in tp else None
        if ps_c is None or ps_c != ps_p or th_c is None or th_c != th_p: kpr += 1
    rows.append((len(kc), nove, izm, kf, kpr))
    print("%-9s %-10s %5d %5d %5d | %6d %6d %6d | %d" % (c[:8], datum, len(kc), nove, izm, kf, kpr, kpr + nove + izm, dirnuti))
def pct(i): return [100.0 * (r[i]) / r[0] for r in rows]
tot = [100.0 * (r[1] + r[2] + r[4]) / r[0] for r in rows]
totf = [100.0 * (r[1] + r[2] + r[3]) / r[0] for r in rows]
print("\nPR-ova: %d" % len(rows))
print("K_proc + nove + izmenjene, %% kataloga po PR-u: medijana %.1f, prosek %.1f, max %.1f" % (statistics.median(tot), sum(tot) / len(tot), max(tot)))
print("K_fajl + nove + izmenjene, %% kataloga po PR-u: medijana %.1f, prosek %.1f, max %.1f" % (statistics.median(totf), sum(totf) / len(totf), max(totf)))
print("zbir (nove+izm+K_proc) preko svih PR-ova: %d ; zbir velicina kataloga: %d" % (sum(r[1] + r[2] + r[4] for r in rows), sum(r[0] for r in rows)))
