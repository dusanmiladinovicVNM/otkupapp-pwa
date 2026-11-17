import re, os, collections, hashlib
A="/tmp/claude-0/-home-user-otkupapp-pwa/d6cefc99-c516-5b26-8cb4-97e3f89ec739/scratchpad/main-wt/src-vba/"
mods=["modTest","modBusinessFlowProTests","modTestBanka","modTestStorno","modTestStornoCentar","modTestPalete","modFakturaTests","modGoldenTests","modIzvestajTests","modAgrohemijaTests","modLicenseTests","modSEFTests","modNovacTests","modMonitoringTests","modGoogleSyncSmokeTests"]
pat=re.compile(r'^(Public |Private |Friend )?(Sub|Function)\s+([A-Za-z_0-9]+)')
byname=collections.defaultdict(list)
for m in mods:
    src=open(A+m+".bas",encoding="latin-1").read().split("\n")
    cur=None
    for i,l in enumerate(src):
        mm=pat.match(l)
        if mm and cur is None: cur=(mm.group(3),i,mm.group(1) or "")
        elif cur and re.match(r'^End (Sub|Function)',l):
            body=[x.strip() for x in src[cur[1]+1:i] if x.strip() and not x.strip().startswith("'")]
            norm="\n".join(body)
            byname[cur[0]].append((m,len(body),hashlib.md5(norm.encode()).hexdigest()[:8],cur[2].strip()))
            cur=None
dups={n:v for n,v in byname.items() if len(v)>1}
print("procedure names defined in >1 test module:",len(dups))
ident=0; lines_dup=0
for n,v in sorted(dups.items(), key=lambda kv:-sum(x[1] for x in kv[1])):
    hs=set(x[2] for x in v)
    same = len(hs)==1
    if same: ident+=1
    lines_dup+=sum(x[1] for x in v[1:])
    print(("IDENT " if same else "DIFF  ")+n, [(x[0].replace("mod",""),x[1]) for x in v])
print("identical bodies:",ident,"/",len(dups),"; redundant body lines (all but first copy):",lines_dup)
