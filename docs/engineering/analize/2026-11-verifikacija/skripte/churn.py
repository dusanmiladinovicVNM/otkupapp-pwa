import subprocess, sys, os, importlib.util, collections
sys.dont_write_bytecode = True
A, S = sys.argv[1], sys.argv[2]
def git(*a):
    return subprocess.run(["git", "-C", A] + list(a), capture_output=True, text=True).stdout
# first-parent linear history of main touching sabotaza.py, oldest first
revs = git("log", "--first-parent", "--format=%H %ad", "--date=short", "--reverse", "--", "tools/sabotaza.py").strip().split("\n")
def load(rev):
    src = git("show", rev + ":tools/sabotaza.py")
    p = os.path.join(S, "hist", "tools", "sab_%s.py" % rev[:10])
    open(p, "w").write(src)
    spec = importlib.util.spec_from_file_location("sab_" + rev[:10], p)
    m = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(m)
    except Exception as e:
        print("  load fail", rev[:8], e); return None
    os.remove(p)
    return dict(m.SABOTAZE)
prev = None; rows = []
tot = collections.Counter()
for line in revs:
    rev, dat = line.split()
    cur = load(rev)
    if cur is None: continue
    if prev is not None:
        add = [k for k in cur if k not in prev]
        rem = [k for k in prev if k not in cur]
        sid = zam = tst = tvr = 0; anyc = 0
        for k in cur:
            if k in prev and cur[k] != prev[k]:
                anyc += 1
                a, b = prev[k], cur[k]
                if a[0] != b[0] or a[1] != b[1]: sid += 1
                if a[2] != b[2]: zam += 1
                if a[3] != b[3]: tst += 1
                if a[4] != b[4]: tvr += 1
        rows.append((dat, rev[:8], len(cur), len(add), len(rem), anyc, sid, zam, tst, tvr))
        tot.update(add=len(add), rem=len(rem), anyc=anyc, sid=sid, zam=zam, tst=tst, tvr=tvr, steps=1)
        if anyc or rem: tot["steps_with_maint"] += 1
        if sid: tot["steps_with_resid"] += 1
    else:
        print("pocetak:", dat, rev[:8], "sabotaza", len(cur))
    prev = cur
print("korak(first-parent commit nad sabotaza.py):", tot["steps"])
print("ukupno: dodato %d, obrisano %d, izmenjenih postojecih %d (sidro/fajl %d, zamena %d, test %d, tvrdnja %d)" % (tot["add"], tot["rem"], tot["anyc"], tot["sid"], tot["zam"], tot["tst"], tot["tvr"]))
print("koraka sa odrzavanjem postojecih (izmena ili brisanje):", tot["steps_with_maint"], "; sa presidrenjem:", tot["steps_with_resid"])
# by week
wk = collections.defaultdict(lambda: collections.Counter())
import datetime
for r in rows:
    d = datetime.date.fromisoformat(r[0]); w = d.isocalendar()[1]
    wk[w].update(add=r[3], rem=r[4], anyc=r[5], sid=r[6], tvr=r[9], n=1); wk[w]["size"] = r[2]
for w in sorted(wk):
    c = wk[w]; print("  ned %d: koraka %d, +%d -%d, izmenjenih %d (sidro %d, tvrdnja %d), katalog na kraju %d" % (w, c["n"], c["add"], c["rem"], c["anyc"], c["sid"], c["tvr"], c["size"]))
print("po datumu (velicina):")
seen = {}
for r in rows: seen[r[0]] = r[2]
for d in sorted(seen): print("  ", d, seen[d])
