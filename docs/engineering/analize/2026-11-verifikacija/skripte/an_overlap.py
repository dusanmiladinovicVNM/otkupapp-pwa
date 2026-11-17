import re, os, collections
A="/tmp/claude-0/-home-user-otkupapp-pwa/d6cefc99-c516-5b26-8cb4-97e3f89ec739/scratchpad/main-wt/src-vba/"
def code(fn):
    out=[]
    for l in open(A+fn,encoding="latin-1").read().split("\n"):
        if l.strip().startswith("'"): continue
        out.append(l)
    return "\n".join(out)
# production public procs
pat=re.compile(r'^(Public )?(Sub|Function)\s+([A-Za-z_0-9]+)',re.M)
prod={}
tests=set(["modTest.bas","modBusinessFlowProTests.bas","modTestBanka.bas","modTestStorno.bas","modTestStornoCentar.bas","modTestPalete.bas","modFakturaTests.bas","modGoldenTests.bas","modIzvestajTests.bas","modAgrohemijaTests.bas","modLicenseTests.bas","modSEFTests.bas","modNovacTests.bas","modMonitoringTests.bas","modGoogleSyncSmokeTests.bas"])
for fn in os.listdir(A):
    if fn.endswith(".bas") and fn not in tests:
        for m in pat.finditer(open(A+fn,encoding="latin-1").read()):
            prod.setdefault(m.group(3),fn)
tok=re.compile(r'[A-Za-z_][A-Za-z_0-9]*')
use={}
for t in ["modTest.bas","modBusinessFlowProTests.bas","modGoldenTests.bas","modTestStorno.bas","modTestPalete.bas","modTestBanka.bas","modTestStornoCentar.bas","modIzvestajTests.bas","modFakturaTests.bas","modNovacTests.bas","modAgrohemijaTests.bas"]:
    c=collections.Counter(x for x in tok.findall(code(t)) if x in prod)
    use[t]=c
a=use["modTest.bas"]; b=use["modBusinessFlowProTests.bas"]
tx=lambda c:{k for k in c if k.endswith("_TX")}
print("prod procs referenced: modTest",len(a),"BFP",len(b),"both",len(set(a)&set(b)))
print("_TX writers: modTest",len(tx(a)),"BFP",len(tx(b)),"both",len(tx(a)&tx(b)))
print("shared _TX:",sorted(tx(a)&tx(b)))
print("only modTest _TX:",sorted(tx(a)-tx(b)))
allTX={k for k in prod if k.endswith("_TX")}
cov=set()
for t,c in use.items(): cov|=tx(c)
print("prod _TX total",len(allTX),"referenced by any test module",len(cov&allTX), "unreferenced:",sorted(allTX-cov)[:40])
for t,c in use.items():
    if t in("modTest.bas","modBusinessFlowProTests.bas"): continue
    print(t,"_TX:",len(tx(c)),"also in modTest:",len(tx(c)&tx(a)),"also in BFP:",len(tx(c)&tx(b)))
# modules exercised
ma=collections.Counter(prod[k] for k in a); mb=collections.Counter(prod[k] for k in b)
print("prod modules touched: modTest",len(ma),"BFP",len(mb),"both",len(set(ma)&set(mb)))
