#!/usr/bin/env python3
"""Po merge PR-u (first-parent origin/main): broj commit-a i promenjene linije po kategoriji. Samo cita git."""
import csv, re, subprocess, os, sys
REPO = "/home/user/otkupapp-pwa"
OUT = os.path.dirname(os.path.abspath(__file__))
def git(*a):
    return subprocess.run(["git", "-C", REPO, *a], capture_output=True, text=True, check=True).stdout

def is_test_module(base):
    if base == "modTestMode.bas": return False
    return bool(re.search(r"Tests?\.bas$", base) or re.match(r"modTest(\w*)\.bas$", base))

def kat(path):
    if path.startswith("src-vba/"):
        return "test_vba" if is_test_module(path.rsplit("/",1)[-1]) else "prod_vba"
    if path == "tools/sabotaza.py": return "sabotaza_py"
    if path.startswith("tools/migracija"): return "tools_migr"
    if re.match(r"tools/(release|stamp-build|logo_to_vba|cenovnik|ugovor|check_merge)", path): return "tools_ost"
    if path.startswith("tools/"): return "tools"
    if path.startswith("tests/"): return "tests_dir"
    if path.startswith(".claude/") or path == "CLAUDE.md": return "claude"
    if path.startswith(".github/"): return "ci"
    if path.startswith("docs/") or path.startswith("backlog/") or path.startswith("instructions/") or path.endswith(".md"): return "docs"
    if path.startswith("src/") or path.startswith("gas/") or path in ("index.html","sw.js","manifest.json","parcel-draw.html"): return "pwa_gas"
    if path.startswith("schema/"): return "schema"
    return "ostalo"
KATS = ["prod_vba","test_vba","sabotaza_py","tools","tools_migr","tools_ost","tests_dir","docs","claude","ci","pwa_gas","schema","ostalo"]

VER_RE = re.compile(r"dokaz|sabota[zž]|sidr[oa]|harness|vba_check|run_vba|fixture|self-?test|\btest|suite|golden|checker|kapij|who_writes|vba_gate|\bgate\b|compile|katalog sabot|dvosmern", re.I)
def klasa_poruke(subj):
    m = re.match(r"^([A-Za-z]+)(\(([^)]*)\))?[:!]", subj)
    pref = (m.group(1).lower() if m else ""); scope = (m.group(3) or "").lower() if m else ""
    if pref in ("docs","spec","process") and not re.search(r"dokaz|sabota|sidr|harness|test|verifik", subj, re.I):
        return "docs"
    if pref in ("test","tests","tools","tool","ci","check","fixture","katalog","modtest","diag","guard") or scope in ("test","tests","dokaz","tools","harness","golden","gate","sabotaza","vba_check","ci","check","fixture","testovi","verifikacija","run_vba","katalog"):
        return "verifikacija"
    if VER_RE.search(subj):
        return "verifikacija" if pref not in ("docs","spec","process") else "docs-o-verifikaciji"
    if pref in ("feat","fix","refactor","perf","harden","revert","style","chore","wip","cleanup","migracija","settings"):
        return "funkcionalno"
    return "ostalo"

since = sys.argv[1] if len(sys.argv) > 1 else "2026-08-09"
merges = [l.split("\t") for l in git("log","--first-parent","--merges","--since="+since,"--format=%H\t%cd\t%s","--date=short","origin/main").splitlines()]
W = csv.writer(open(os.path.join(OUT,"pr.csv"),"w",newline=""))
W.writerow(["dan","merge","pr","grana","commit_a","c_verif","c_docs","c_docs_o_verif","c_funkc","c_ostalo"]+[k+"_add" for k in KATS]+[k+"_del" for k in KATS])
for h,d,s in reversed(merges):
    m = re.search(r"#(\d+) from \S+?/(\S+)", s)
    pr = m.group(1) if m else ""; grana = m.group(2) if m else s[:40]
    subs = git("log","--no-merges","--format=%s",f"{h}^1..{h}^2").splitlines()
    kl = {"verifikacija":0,"docs":0,"docs-o-verifikaciji":0,"funkcionalno":0,"ostalo":0}
    for x in subs: kl[klasa_poruke(x)] += 1
    add = dict.fromkeys(KATS,0); dele = dict.fromkeys(KATS,0)
    for l in git("diff","--numstat","--no-renames",f"{h}^1",h).splitlines():
        a,b,p = l.split("\t",2)
        if a == "-": continue
        k = kat(p); add[k]+=int(a); dele[k]+=int(b)
    W.writerow([d,h[:8],pr,grana,len(subs),kl["verifikacija"],kl["docs"],kl["docs-o-verifikaciji"],kl["funkcionalno"],kl["ostalo"]]+[add[k] for k in KATS]+[dele[k] for k in KATS])
print("merge PR-ova:", len(merges))
