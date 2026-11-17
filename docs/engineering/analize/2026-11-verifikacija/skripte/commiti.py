#!/usr/bin/env python3
"""Po commit-u (bez merge-a) od --since: klasa po fajlovima + klasa po poruci. Samo cita git."""
import csv, re, subprocess, os, sys, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
REPO = "/home/user/otkupapp-pwa"
OUT = os.path.dirname(os.path.abspath(__file__))
import importlib.util
src = open(os.path.join(OUT, "pr.py")).read().split("since = sys.argv")[0]
ns = {"__file__": os.path.join(OUT, "pr.py")}; exec(src, ns)
kat = ns["kat"]; klasa_poruke = ns["klasa_poruke"]; KATS = ns["KATS"]
since = sys.argv[1] if len(sys.argv) > 1 else "2026-08-09"
out = subprocess.run(["git","-C",REPO,"log","--no-merges","--since="+since,"--format=@@%H\t%cd\t%s","--date=short","--numstat","--no-renames","origin/main"],capture_output=True,text=True,check=True).stdout
VER=("test_vba","sabotaza_py","tools","tests_dir","ci"); DOC=("docs","claude"); PROD=("prod_vba","pwa_gas","schema","tools_migr","tools_ost")
W = csv.writer(open(os.path.join(OUT,"commiti.csv"),"w",newline=""))
W.writerow(["dan","commit","klasa_poruke","klasa_fajlova","subject"]+KATS)
cur=None; rows=[]
def flush():
    if not cur: return
    h,d,s,ch = cur
    v = sum(ch[k] for k in VER); dc = sum(ch[k] for k in DOC); p = sum(ch[k] for k in PROD)
    if p and v: kf = "prod+verif"
    elif p: kf = "samo-prod"
    elif v: kf = "samo-verif"
    elif dc: kf = "samo-docs"
    else: kf = "prazno/ostalo"
    W.writerow([d,h[:8],klasa_poruke(s),kf,s]+[ch[k] for k in KATS])
for l in out.splitlines():
    if l.startswith("@@"):
        flush()
        h,d,s = l[2:].split("\t",2)
        cur=(h,d,s,dict.fromkeys(KATS,0))
    elif l.strip():
        a,b,p = l.split("\t",2)
        if a=="-": continue
        cur[3][kat(p)] += int(a)+int(b)
flush()
