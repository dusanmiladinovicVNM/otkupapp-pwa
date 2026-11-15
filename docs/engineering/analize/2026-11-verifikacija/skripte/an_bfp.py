import re, sys, json, collections
A="/tmp/claude-0/-home-user-otkupapp-pwa/d6cefc99-c516-5b26-8cb4-97e3f89ec739/scratchpad/main-wt/src-vba/"
src=open(A+"modBusinessFlowProTests.bas",encoding="latin-1").read().split("\n")
pat=re.compile(r'^(Public |Private |Friend )?(Sub|Function)\s+([A-Za-z_0-9]+)')
procs={}; cur=None
for i,l in enumerate(src):
    m=pat.match(l)
    if m and cur is None: cur=(m.group(3),i)
    elif cur and re.match(r'^End (Sub|Function)',l):
        procs[cur[0]]=(cur[1],i); cur=None
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
s,e=procs["RunBusinessFlowProSuite"]
order=[]
for l in src[s+1:e]:
    t=l.strip()
    if t.startswith("'") or not t: continue
    m=re.match(r'(Test_[A-Za-z0-9_]+|Seed[A-Za-z0-9_]+)\s*$',t)
    if m: order.append(m.group(1))
print("procs",len(procs),"tests in suite",len(order),"unique",len(set(order)), "Test_ procs total",sum(1 for n in procs if n.startswith("Test_")))
print("Test_ procs not in suite:",sorted(n for n in procs if n.startswith("Test_") and n not in order)[:20])
RX={"begintx": r'\.BeginTx\b',"rollback": r'\.RollbackTx\b',"commit": r'\.CommitTx\b',"snap": r'AddTableSnapshot',
 "prodTX": r'\b[A-Za-z0-9]+[A-Za-z0-9_]*_TX\b',"append": r'\bAppendRow\b|RequireAppend',"update": r'\bUpdateCell\b|RequireUpdateCell',
 "delete": r'\bDeleteRow\b|ListRows\([^)]*\)\.Delete|\.Delete\b',"config": r'\bSetConfigValue\b',"msgbox": r'\bMsgBox\b',
 "newscn": r'NewScenarioCode',"nextdate": r'NextTestDate',"assert": r'\b(AssertTrue|AssertFalse|AssertEquals|AssertDoubleNear|LogPass|LogFail)\b',
 "gettabledata": r'\bGetTableData\b', "form": r'New frm', "softstorno": r'MarkTestRowStornirano|Storno[A-Za-z]*_TX|StornoDokument'}
rows=[]
helper_use=collections.Counter()
for pos,n in enumerate(order):
    cl=closure(n)|{n}
    txt="\n".join(bodies[x] for x in cl)
    own=bodies[n]
    r={"pos":pos+1,"name":n,"lines":procs[n][1]-procs[n][0]+1}
    for k,rx in RX.items():
        r[k]=len(re.findall(rx,txt)); r[k+"_own"]=len(re.findall(rx,own))
    for h in cl-{n}: helper_use[h]+=1
    mut = r["prodTX"] or r["append"] or r["update"] or r["delete"]
    if not mut and not r["config"]: r["cls"]="A_citanje"
    elif not mut: r["cls"]="C0_config"
    elif r["begintx_own"] and r["rollback_own"] and r["snap_own"]: r["cls"]="C_snimak+rollback(u testu)"
    elif r["begintx"] and r["rollback"] and r["snap"]: r["cls"]="C2_snimak+rollback(u helperu)"
    else: r["cls"]="D_upis_bez_rollbacka"
    rows.append(r)
cnt=collections.Counter(r["cls"] for r in rows)
for k,v in sorted(cnt.items()): print(v,k)
print("asserts (static, own body) sum:",sum(r["assert_own"] for r in rows), "lines sum",sum(r["lines"] for r in rows))
print("NewScenarioCode users:",sum(1 for r in rows if r["newscn"]), "NextTestDate users:",sum(1 for r in rows if r["nextdate"]))
print("tests with SetConfigValue:",sum(1 for r in rows if r["config"]))
print("top helpers:",helper_use.most_common(40))
json.dump(rows,open(sys.argv[1],"w"))
print("seq:","".join({"A":"r","C0":"c","C_":"T","C2":"t","D_":"M"}[r["cls"][:2]] for r in rows))
