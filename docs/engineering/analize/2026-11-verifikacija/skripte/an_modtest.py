import re, sys, json, collections
A="/tmp/claude-0/-home-user-otkupapp-pwa/d6cefc99-c516-5b26-8cb4-97e3f89ec739/scratchpad/main-wt/src-vba/"
src=open(A+"modTest.bas",encoding="latin-1").read().split("\n")
# procedures
procs={}  # name -> (start,end,kind)
cur=None
pat=re.compile(r'^(Public|Private)?\s*(Sub|Function|Property Get|Property Let)\s+([A-Za-z_0-9]+)')
for i,l in enumerate(src):
    m=pat.match(l)
    if m and cur is None:
        cur=(m.group(3),i)
    elif cur and re.match(r'^End (Sub|Function|Property)',l):
        procs[cur[0]]=(cur[1],i)
        cur=None
# registers
def reg(procname, rx):
    s,e=procs[procname]
    d={}
    for l in src[s:e]:
        m=re.match(rx,l.strip())
        if m: d[int(m.group(1))]=m.group(2)
    return d
inv=reg("InvokeTest", r'Case (\d+): (T_[A-Za-z0-9_]+)')
nm=reg("TestName", r'Case (\d+): TestName = "(T_[A-Za-z0-9_]+)"')
s,e=procs["RunAllTests"]
order=[int(m.group(1)) for l in src[s:e] for m in [re.match(r'\s*RunOne (\d+)',l)] if m]
print("InvokeTest:",len(inv),"TestName:",len(nm),"RunOne:",len(order),"unique",len(set(order)))
print("mismatch name:",[k for k in inv if nm.get(k)!=inv[k]][:10])
print("not in order:",sorted(set(inv)-set(order)), "order not in inv:",sorted(set(order)-set(inv)))
# helper call graph: body text per proc (without comments)
def body(name):
    s,e=procs[name]
    out=[]
    for l in src[s+1:e]:
        t=l.strip()
        if t.startswith("'"): continue
        out.append(l)
    return "\n".join(out)
bodies={n:body(n) for n in procs}
names=set(procs)
calls={}
tok=re.compile(r'[A-Za-z_][A-Za-z_0-9]*')
for n,b in bodies.items():
    calls[n]=set(t for t in tok.findall(b) if t in names and t!=n)
def closure(n):
    seen=set(); st=[n]
    while st:
        x=st.pop()
        for c in calls.get(x,()):
            if c not in seen:
                seen.add(c); st.append(c)
    return seen
json.dump({"procs":procs,"inv":inv,"order":order}, open(sys.argv[1],"w"))
# classification
KW={
 "form": r'\bNew frmOtkupUI\b|NapraviOtkupUI|OtkupUIForm|NovaForma|BuildOtkupUI|OtkupUI_Init|OtkupUI_Bind|MakeOtkupUIForm|NapraviFormu',
 "tx_begin": r'\.BeginTx\b',
 "tx_rb": r'\.RollbackTx\b',
 "tx_commit": r'\.CommitTx\b',
 "writer_TX": r'\b[A-Za-z0-9_]+_TX\b',
 "append": r'\bAppendRow\b|ListRows\.Add',
 "update": r'\bUpdateCell\b|UpdateRow|\.Value2? *=|\.Range\(.*\)\.Value *=',
 "delete": r'ListRows\([^)]*\)\.Delete|DeleteRow|\.Delete\b',
 "invalidate": r'InvalidateTableCache|InvalidateCache',
 "config": r'SetConfigValue|SetConfig\b|ConfigSet',
 "file": r'\bOpen .* For |Kill |MkDir|WriteTextFile',
 "sheet": r'Worksheets\.Add|Sheets\.Add',
 "screenupd": r'ScreenUpdating',
 "calc": r'Application\.Calculation',
}
rows=[]
for idx in order:
    n=inv[idx]
    cl=closure(n)|{n}
    txt="\n".join(bodies[x] for x in cl)
    own=bodies[n]
    r={"idx":idx,"name":n,"lines":procs[n][1]-procs[n][0]+1,
       "asserts_own":len(re.findall(r'\bAssertEq\b|\bAssertSnapshot\b',own)),
       "asserts_cl":len(re.findall(r'\bAssertEq\b|\bAssertSnapshot\b',txt))}
    for k,rx in KW.items():
        r[k]=len(re.findall(rx,txt))
        r[k+"_own"]=len(re.findall(rx,own))
    rows.append(r)
json.dump(rows, open(sys.argv[2],"w"))
print("tests total lines:",sum(r["lines"] for r in rows), "asserts own total:",sum(r["asserts_own"] for r in rows))
import statistics
print("lines/test median",statistics.median(r["lines"] for r in rows),"max",max((r["lines"],r["name"]) for r in rows))
for k in KW:
    print(k, sum(1 for r in rows if r[k]>0), "own:",sum(1 for r in rows if r[k+"_own"]>0))
