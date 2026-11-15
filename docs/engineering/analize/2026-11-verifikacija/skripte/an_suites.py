import re, sys, json
A="/tmp/claude-0/-home-user-otkupapp-pwa/d6cefc99-c516-5b26-8cb4-97e3f89ec739/scratchpad/main-wt/src-vba/"
SU=[("RunIzvestajTests","modIzvestajTests"),("RunSheetsJsonParserTests","modGoogleSyncSmokeTests"),
("RunBankaImportTestSuite","modTestBanka"),("RunFakturaSmokeSuite","modFakturaTests"),("Test_StornoCentar_All","modTestStornoCentar"),
("TestLicense_All","modLicenseTests"),("RunGoogleSyncSmokeSuite","modGoogleSyncSmokeTests"),("RunMasterSyncSmokeSuite","modGoogleSyncSmokeTests"),
("RunSEFTestSuite","modSEFTests"),("RunStornoTestSuite","modTestStorno"),("RunPaleteTestSuite","modTestPalete"),("RunNovacSmokeSuite","modNovacTests"),
("RunBusinessFlowProSuite","modBusinessFlowProTests"),("RunAgrohemijaSmokeSuite","modAgrohemijaTests"),("RunGoldenSuite","modGoldenTests"),
("RunProductionHealthCheck","modProductionHealthCheck"),("TestMonitoring_All","modMonitoringTests")]
pat=re.compile(r'^(Public |Private |Friend )?(Sub|Function)\s+([A-Za-z_0-9]+)')
tok=re.compile(r'[A-Za-z_][A-Za-z_0-9]*')
def load(mod):
    src=open(A+mod+".bas",encoding="latin-1").read().split("\n")
    procs={}; cur=None
    for i,l in enumerate(src):
        m=pat.match(l)
        if m and cur is None: cur=(m.group(3),i)
        elif cur and re.match(r'^End (Sub|Function)',l):
            procs[cur[0]]=(cur[1],i); cur=None
    bodies={}
    for n,(s,e) in procs.items():
        bodies[n]="\n".join(l for l in src[s+1:e] if not l.strip().startswith("'"))
    return src,procs,bodies
ASSERT=r'\b(AssertEq|AssertTrue|AssertFalse|AssertEquals|AssertDoubleNear|AssertSnapshot|Check|Expect[A-Za-z]*|Assert[A-Za-z]*|LogPass|LogFail|TPass|TFail|Pass|Fail|Tvrdi[A-Za-z]*|Ok|Proveri[A-Za-z]*)\b'
for suite,mod in SU:
    src,procs,bodies=load(mod)
    names=set(procs)
    calls={n:[t for t in tok.findall(b) if t in names and t!=n] for n,b in bodies.items()}
    if suite not in procs: print(suite,"NOT FOUND"); continue
    direct=[]
    for t in calls[suite]:
        if t not in direct: direct.append(t)
    seen=set(); st=[suite]
    while st:
        x=st.pop()
        for c in calls.get(x,()):
            if c not in seen: seen.add(c); st.append(c)
    txt="\n".join(bodies[x] for x in seen|{suite})
    own=bodies[suite]
    testlike=[d for d in direct if re.match(r'(Test|T_|TC_|Smoke|Golden|G_|Scn|Scenario|Run)',d)]
    print("==",suite,mod,"lines",procs[suite][1]-procs[suite][0]+1,"direct calls",len(direct),"test-like",len(testlike),
          "| MsgBox in closure",len(re.findall(r'\bMsgBox\b',txt)),"own",len(re.findall(r'\bMsgBox\b',own)),
          "| BeginTx",len(re.findall(r'\.BeginTx\b',txt)),"Rollback",len(re.findall(r'\.RollbackTx\b',txt)),"Commit",len(re.findall(r'\.CommitTx\b',txt)),
          "| _TX calls",len(re.findall(r'\b[A-Za-z0-9_]+_TX\b',txt)),"AppendRow",len(re.findall(r'\bAppendRow\b',txt)),"UpdateCell",len(re.findall(r'UpdateCell\b',txt)),
          "DeleteRow/.Delete",len(re.findall(r'DeleteRow\b|\.Delete\b',txt)),
          "| Err.Raise",len(re.findall(r'Err\.Raise',txt)),"| asserts~",len(re.findall(r'\b(AssertEq|AssertTrue|AssertFalse|AssertEquals|AssertDoubleNear|AssertSnapshot)\b',txt)))
    print("   direct:",", ".join(direct[:60]))
