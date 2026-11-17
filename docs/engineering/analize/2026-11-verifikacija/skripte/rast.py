#!/usr/bin/env python3
"""Krive rasta po danu nad first-parent istorijom origin/main. Samo cita git."""
import ast, csv, re, subprocess, sys, os
REPO = "/home/user/otkupapp-pwa"
OUT = os.path.dirname(os.path.abspath(__file__))

def git(*a, text=True):
    return subprocess.run(["git", "-C", REPO, *a], capture_output=True, text=text, check=True).stdout

# first-parent commits, newest first
rows = [l.split("\t") for l in git("log", "--first-parent", "--format=%H\t%cd", "--date=short", "origin/main").splitlines()]
last_of_day = {}
for h, d in rows:
    last_of_day.setdefault(d, h)   # newest-first => first seen is last of the day
days = sorted(last_of_day)

cat = subprocess.Popen(["git", "-C", REPO, "cat-file", "--batch"], stdin=subprocess.PIPE, stdout=subprocess.PIPE)
blob_cache = {}
def blob(sha):
    if sha in blob_cache:
        return blob_cache[sha]
    cat.stdin.write((sha + "\n").encode()); cat.stdin.flush()
    hdr = cat.stdout.readline().split()
    size = int(hdr[2])
    data = cat.stdout.read(size); cat.stdout.read(1)
    blob_cache[sha] = data
    return data
lines_cache = {}
def nlines(sha):
    if sha not in lines_cache:
        b = blob(sha)
        lines_cache[sha] = b.count(b"\n") + (1 if b and not b.endswith(b"\n") else 0)
    return lines_cache[sha]

TEST_RE = re.compile(r"(Tests?\.bas$|/modTest[A-Z]\w*\.bas$|/modTest\.bas$)")
def is_test_module(path):
    base = path.rsplit("/", 1)[-1]
    if base == "modTestMode.bas":
        return False       # produkcioni seam (mora postojati i u produkcijskom build-u)
    return bool(re.search(r"Tests?\.bas$", base) or re.match(r"modTest(\w*)\.bas$", base))

def n_sabotaza(src):
    try:
        tree = ast.parse(src)
    except SyntaxError:
        return -1
    n = 0
    for node in tree.body:
        tgt = None
        if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            tgt = node.targets[0].id; val = node.value
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            tgt = node.target.id; val = node.value
        if tgt == "SABOTAZE" and isinstance(val, ast.Dict):
            n += len(val.keys)
    return n

def sab_po_suiti(src):
    """Broj sabotaza po fajlu testa (5. element tuple-a pocinje sa T_ / Test_)."""
    try:
        tree = ast.parse(src)
    except SyntaxError:
        return {}
    out = {"T_": 0, "Test_": 0, "drugo": 0, "fajlova": set(), "testova": set()}
    for node in tree.body:
        if isinstance(node, ast.Assign) and isinstance(node.targets[0], ast.Name) and node.targets[0].id == "SABOTAZE" and isinstance(node.value, ast.Dict):
            for v in node.value.values:
                if isinstance(v, ast.Tuple) and len(v.elts) >= 4:
                    f = v.elts[0]; t = v.elts[3]
                    if isinstance(f, ast.Constant): out["fajlova"].add(f.value)
                    if isinstance(t, ast.Constant) and isinstance(t.value, str):
                        out["testova"].add(t.value)
                        if t.value.startswith("T_"): out["T_"] += 1
                        elif t.value.startswith("Test_"): out["Test_"] += 1
                        else: out["drugo"] += 1
    return out

def bfp_calls(txt):
    m = re.search(r"^Public Sub RunBusinessFlowProSuite\(\)(.*?)^End Sub", txt, re.S | re.M)
    if not m:
        return 0
    return len(re.findall(r"^[ \t]+Test_\w+[ \t]*\r?$", m.group(1), re.M))

W = csv.writer(open(os.path.join(OUT, "rast.csv"), "w", newline=""))
hdr = ["dan", "commit", "sabotaza", "sab_T", "sab_Test", "sab_fajlova", "sab_testova", "runone", "T_def", "bfp_poziva", "bfp_Test_def",
       "modTest_lin", "bfp_lin", "test_mod_lin", "test_mod_n", "prod_vba_lin", "prod_vba_n", "srcvba_fajlova",
       "tools_py_lin", "tools_py_n", "sabotaza_py_lin", "vba_check_lin", "run_vba_lin", "dokaz_lin", "docs_md_lin", "docs_md_n", "rules_lin"]
W.writerow(hdr)
for d in days:
    c = last_of_day[d]
    ls = git("ls-tree", "-r", c, "--", "src-vba", "tools", "docs", ".claude", "CLAUDE.md")
    files = {}
    for l in ls.splitlines():
        meta, path = l.split("\t", 1)
        mode, typ, sha = meta.split()
        if typ == "blob":
            files[path] = sha
    vba = {p: s for p, s in files.items() if p.startswith("src-vba/")}
    vba_text = {p: s for p, s in vba.items() if p.rsplit(".", 1)[-1].lower() in ("bas", "cls", "doccls", "frm")}
    test_mod = {p: s for p, s in vba_text.items() if is_test_module(p)}
    prod = {p: s for p, s in vba_text.items() if not is_test_module(p)}
    tools_py = {p: s for p, s in files.items() if re.match(r"tools/[^/]+\.py$", p)}
    docs_md = {p: s for p, s in files.items() if p.startswith("docs/") and p.endswith(".md")}
    rules = {p: s for p, s in files.items() if p.startswith(".claude/") or p == "CLAUDE.md"}
    def L(p):
        return nlines(files[p]) if p in files else 0
    sab = sabT = sabTest = sabF = sabN = 0
    if "tools/sabotaza.py" in files:
        src = blob(files["tools/sabotaza.py"]).decode("utf-8", "replace")
        sab = n_sabotaza(src)
        ps = sab_po_suiti(src)
        if ps:
            sabT, sabTest, sabF, sabN = ps["T_"], ps["Test_"], len(ps["fajlova"]), len(ps["testova"])
    runone = tdef = 0
    if "src-vba/modTest.bas" in files:
        t = blob(files["src-vba/modTest.bas"]).decode("latin-1")
        runone = len(re.findall(r"^[ \t]*RunOne[ \t]+\d+", t, re.M))
        tdef = len(re.findall(r"^(?:Private|Public) (?:Sub|Function) T_\w+", t, re.M))
    bc = bd = 0
    if "src-vba/modBusinessFlowProTests.bas" in files:
        t = blob(files["src-vba/modBusinessFlowProTests.bas"]).decode("latin-1")
        bc = bfp_calls(t)
        bd = len(re.findall(r"^(?:Private|Public) (?:Sub|Function) Test_\w+", t, re.M))
    W.writerow([d, c[:8], sab, sabT, sabTest, sabF, sabN, runone, tdef, bc, bd,
                L("src-vba/modTest.bas"), L("src-vba/modBusinessFlowProTests.bas"),
                sum(nlines(s) for s in test_mod.values()), len(test_mod),
                sum(nlines(s) for s in prod.values()), len(prod), len(vba),
                sum(nlines(s) for s in tools_py.values()), len(tools_py),
                L("tools/sabotaza.py"), L("tools/vba_check.py"), L("tools/run_vba.py"), L("tools/dokaz.py"),
                sum(nlines(s) for s in docs_md.values()), len(docs_md),
                sum(nlines(s) for s in rules.values())])
cat.stdin.close()
print("dana:", len(days), "prvi:", days[0], "poslednji:", days[-1])
