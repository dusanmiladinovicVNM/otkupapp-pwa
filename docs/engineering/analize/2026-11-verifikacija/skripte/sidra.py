#!/usr/bin/env python3
"""Po merge PR-u: koliko je POSTOJECIH unosa kataloga sabotaza izmenjeno (presidreno), obrisano, dodato. Samo cita git (ast, bez izvrsavanja)."""
import ast, csv, re, subprocess, os, sys
REPO = "/home/user/otkupapp-pwa"
OUT = os.path.dirname(os.path.abspath(__file__))
def git(*a, check=True):
    r = subprocess.run(["git","-C",REPO,*a],capture_output=True,text=True)
    return r.stdout if r.returncode==0 else None
cache={}
def katalog(rev):
    sha = git("rev-parse", rev+":tools/sabotaza.py")
    if sha is None: return {}
    sha=sha.strip()
    if sha in cache: return cache[sha]
    src = git("cat-file","-p",sha)
    out={}
    try:
        tree=ast.parse(src)
        for node in tree.body:
            if isinstance(node,ast.Assign) and isinstance(node.targets[0],ast.Name) and node.targets[0].id=="SABOTAZE" and isinstance(node.value,ast.Dict):
                for k,v in zip(node.value.keys,node.value.values):
                    try: out[ast.literal_eval(k)] = ast.literal_eval(v)
                    except Exception: out[ast.unparse(k)] = ("?",ast.unparse(v))
    except SyntaxError: pass
    cache[sha]=out
    return out
since = sys.argv[1] if len(sys.argv)>1 else "2026-08-09"
merges=[l.split("\t") for l in git("log","--first-parent","--merges","--since="+since,"--reverse","--format=%H\t%cd\t%s","--date=short","origin/main").splitlines()]
W=csv.writer(open(os.path.join(OUT,"sidra.csv"),"w",newline=""))
W.writerow(["dan","merge","pr","pre","posle","dodato","obrisano","izmenjeno","izm_sidro","izm_zamena","izm_test","izm_tvrdnja","izm_fajl"])
for h,d,s in merges:
    a=katalog(h+"^1"); b=katalog(h)
    if not a and not b: continue
    add=[k for k in b if k not in a]; rem=[k for k in a if k not in b]
    ch=[k for k in b if k in a and a[k]!=b[k]]
    def dif(i): return sum(1 for k in ch if len(a[k])>i and len(b[k])>i and a[k][i]!=b[k][i])
    m=re.search(r"#(\d+)",s)
    W.writerow([d,h[:8],m.group(1) if m else "",len(a),len(b),len(add),len(rem),len(ch),dif(1),dif(2),dif(3),dif(4),dif(0)])
