#!/usr/bin/env python3
"""Sva zabelezena trajanja/rezultati: poruke commit-a + DODATE linije u docs/, .claude/, CLAUDE.md, tools/ komentari.
Datum = datum commit-a koji je liniju dodao. Samo cita git."""
import re, subprocess, os, sys, csv
REPO = "/home/user/otkupapp-pwa"
OUT = os.path.dirname(os.path.abspath(__file__))
since = sys.argv[1] if len(sys.argv) > 1 else "2026-08-01"
log = subprocess.run(["git","-C",REPO,"log","--no-merges","--since="+since,"--reverse","--format=\x01%H\t%cd\t%s\x02%b\x03","--date=short","-p","-U0","--no-renames","origin/main","--","docs",".claude","CLAUDE.md"],capture_output=True,text=True,errors="replace").stdout
# commit-i koji ne diraju docs nemaju patch ovde; poruke se uzimaju posebno
PAT = [
 ("suite_s",   re.compile(r"(Run\w+|BFP|Palete|Storno|Golden|banka)\w*\s*\(\s*(\d+(?:[.,]\d+)?)\s*s\)")),
 ("s_blizu",   re.compile(r"(RunAllTests|RunBusinessFlowProSuite|BFP|FULL|dokaz\w*|prolaz\w*|suite\w*|Compile|rezij\w*|re[zž]ij\w*)[^\n]{0,60}?(?<![\w.,])(\d+(?:[.,]\d+)?)\s*(s|sek\w*)\b", re.I)),
 ("min",       re.compile(r"(?<![\w.,])(~?\d+(?:[.,]\d+)?(?:\s*[-–]\s*\d+)?)\s*(min\w*)\b", re.I)),
 ("sat",       re.compile(r"(?<![\w.,])(\d+(?:[.,]\d+)?|dva i po|pola|jedan|dva|tri)\s*(sat|sata|sati|h)\b", re.I)),
 ("ukupno_palo", re.compile(r"(\d+) ukupno, (\d+) palo")),
 ("rat_n",     re.compile(r"RunAllTests\*{0,2}[^0-9\n]{0,14}?\*{0,2}(\d+)\*{0,2}\s*/\s*(\d+)")),
 ("bfp_n",     re.compile(r"(?:BFP|RunBusinessFlowProSuite)`?\*{0,2}[^0-9\n]{0,16}?\*{0,2}(\d+)\*{0,2}\s*/\s*(\d+)")),
 ("dokaz_n",   re.compile(r"dokaz\w*[^0-9\n]{0,30}?\*{0,2}(\d+)\*{0,2}\s*/\s*(\d+)", re.I)),
 ("prolaz_traje", re.compile(r"prolaz traje[^\n]{0,60}", re.I)),
]
W = csv.writer(open(os.path.join(OUT,"trajanja_raw.csv"),"w",newline=""))
W.writerow(["dan","commit","izvor","vrsta","pogodak","kontekst"])
n=0
for blok in log.split("\x01")[1:]:
    head, rest = blok.split("\x02",1)
    h,d,s = head.split("\t",2)
    body, patch = rest.split("\x03",1)
    izvori = [("poruka", s+"\n"+body)]
    cur=None
    for l in patch.splitlines():
        if l.startswith("+++ b/"): cur=l[6:]
        elif l.startswith("+") and not l.startswith("+++") and cur:
            izvori.append((cur, l[1:]))
    for izv, txt in izvori:
        for vrsta, rx in PAT:
            for m in rx.finditer(txt):
                a=max(0,m.start()-110); b=min(len(txt),m.end()+90)
                W.writerow([d,h[:8],izv,vrsta,m.group(0),txt[a:b].replace("\n"," ")]); n+=1
print("pogodaka:", n)
