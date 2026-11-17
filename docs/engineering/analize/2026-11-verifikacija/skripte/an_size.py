import re, os, sys, collections
A="/tmp/claude-0/-home-user-otkupapp-pwa/d6cefc99-c516-5b26-8cb4-97e3f89ec739/scratchpad/main-wt/src-vba/"
pat=re.compile(r'^(Public |Private |Friend )?(Static )?(Sub|Function|Property (Get|Let|Set))\s+([A-Za-z_0-9]+)')
allp=[]
ids=set()
tokre=re.compile(r'[A-Za-z_][A-Za-z_0-9]*')
strre=re.compile(r'"(?:[^"]|"")*"')
tot_lines=0
per_mod={}
for fn in sorted(os.listdir(A)):
    if not fn.endswith((".bas",".cls",".frm",".doccls")): continue
    src=open(A+fn,encoding="latin-1").read().split("\n")
    tot_lines+=len(src)
    cur=None; code=0; nl=0; ncomment=0
    for l in src:
        t=l.strip()
        if t.startswith("'"): ncomment+=1
        s=strre.sub('""',l)
        if "'" in s: s=s[:s.index("'")]
        for tk in tokre.findall(s): ids.add(tk.lower())
        m=pat.match(l)
        if m and cur is None:
            cur=m.group(5); code=0; nl=0
        if cur is not None:
            nl+=1
            if not t.startswith("'"): code+=len(s.strip())
            if re.match(r'^End (Sub|Function|Property)',l):
                allp.append((code,nl,fn,cur)); cur=None
    per_mod[fn]=(len(src),ncomment)
allp.sort(reverse=True)
print("unique identifiers (case-insens., all src-vba):",len(ids),"total lines",tot_lines)
print("top procedures by code chars (comments stripped):")
for c,n,f,p in allp[:15]: print(c,n,f,p)
for f in ("modTest.bas","modBusinessFlowProTests.bas"):
    L,C=per_mod[f]; print(f,"lines",L,"comment lines",C,"=%.0f%%"%(100*C/L), "procs",sum(1 for x in allp if x[2]==f))
tp=[x for x in allp if x[2] in("modTest.bas","modBusinessFlowProTests.bas")]
print("top test procs:"); 
for c,n,f,p in tp[:8]: print(c,n,f,p)
