import re, sys, json, collections
A="/tmp/claude-0/-home-user-otkupapp-pwa/d6cefc99-c516-5b26-8cb4-97e3f89ec739/scratchpad/main-wt/src-vba/"
src=open(A+"modTest.bas",encoding="latin-1").read().split("\n")
d=json.load(open(sys.argv[1]))
procs=d["procs"]; inv={int(k):v for k,v in d["inv"].items()}; order=d["order"]
def body(n):
    s,e=procs[n]
    return "\n".join(l for l in src[s+1:e] if not l.strip().startswith("'"))
bodies={n:body(n) for n in procs}
tok=re.compile(r'[A-Za-z_][A-Za-z_0-9]*')
names=set(procs)
calls={n:set(t for t in tok.findall(b) if t in names and t!=n) for n,b in bodies.items()}
def closure(n):
    seen=set(); st=[n]
    while st:
        x=st.pop()
        for c in calls.get(x,()):
            if c not in seen: seen.add(c); st.append(c)
    return seen
RX={
 "form": r'\bNewOtkupUIForm\b|\bNew frmOtkupUI\b',
 "begintx": r'\.BeginTx\b',
 "rollback": r'\.RollbackTx\b',
 "commit": r'\.CommitTx\b',
 "snap": r'AddTableSnapshot',
 "prodTX": r'\b(?!T_)[A-Za-z0-9]+[A-Za-z0-9_]*_TX\b',
 "append": r'\bAppendRow\b',
 "update": r'\bUpdateCell\b|RequireUpdateCell',
 "delete": r'\bDeleteRow\b|ListRows\([^)]*\)\.Delete|\.Delete\b',
 "config": r'\bSetConfigValue\b',
 "storno": r'\bStorno[A-Za-z_]*\(|\bStorno[A-Za-z_]* ',
 "cellwrite": r'\.(Value|Value2|Formula|NumberFormat)\s*=\s',
 "auth": r'AuthTestUkljuci|AuthTest',
 "sheetadd": r'Worksheets\.Add|TxRbNapraviList',
 "file": r'\bKill\b|Open .* For (Output|Append)|MkDir|FileCopy',
}
rows=[]
for pos,idx in enumerate(order):
    n=inv[idx]; cl=closure(n)|{n}
    txt="\n".join(bodies[x] for x in cl)
    r={"pos":pos+1,"idx":idx,"name":n,"lines":procs[n][1]-procs[n][0]+1}
    for k,rx in RX.items(): r[k]=len(re.findall(rx,txt))
    mut = r["prodTX"] or r["append"] or r["update"] or r["delete"] or r["storno"]
    if not mut and not r["config"] and not r["cellwrite"]:
        r["cls"]="B_forma_bez_upisa" if r["form"] else "A_citanje"
    elif not mut and (r["config"] or r["cellwrite"]):
        r["cls"]="C0_config_ili_celija" + ("+forma" if r["form"] else "")
    else:
        if r["begintx"] and r["rollback"] and r["snap"]:
            r["cls"]="C_upis_sa_snimkom_i_rollback" + ("+forma" if r["form"] else "")
        else:
            r["cls"]="D_upis_bez_test_rollbacka" + ("+forma" if r["form"] else "")
    rows.append(r)
cnt=collections.Counter(r["cls"] for r in rows)
for k,v in sorted(cnt.items()): print(v,k)
print("forma ukupno:",sum(1 for r in rows if r["form"]))
json.dump(rows,open(sys.argv[2],"w"))
print()
for r in rows:
    if r["cls"].startswith("D"):
        print(r["pos"],r["idx"],r["name"],"TX",r["prodTX"],"app",r["append"],"upd",r["update"],"del",r["delete"],"storno",r["storno"],"cfg",r["config"],"begin",r["begintx"],"rb",r["rollback"],"snap",r["snap"])
