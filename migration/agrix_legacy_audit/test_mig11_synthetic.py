import json
import tempfile
from pathlib import Path
import importlib.util
from contextlib import redirect_stdout, redirect_stderr
from io import StringIO

p=Path('/mnt/data/agrix_mig11_build_links.py')
spec=importlib.util.spec_from_file_location('migration',p)
m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)

def save(path,data):
    path.write_text(''.join(json.dumps(x,ensure_ascii=False)+'\n' for x in data),encoding='utf-8')

def fixture(base):
    base.mkdir()
    canonO=[{'OtkupID':'OTK-1','StanicaID':'ST-1','Stornirano':'Ne'},
            {'OtkupID':'OTK-2','StanicaID':'ST-1','Stornirano':'Ne'},
            {'OtkupID':'OTK-3','StanicaID':'ST-1','Stornirano':'Da'},
            {'OtkupID':'OTK-4','StanicaID':'ST-1','Stornirano':'Ne'}]
    srcO=[{**r,'OtpremnicaID':f'OTP-{i}' if i in (1,2,3) else ''} for i,r in enumerate(canonO,1)]
    canonP=[{'OtpremnicaID':'OTP-1','StanicaID':'ST-1','BrojOtpremnice':'101','Stornirano':'Ne','VozacID':'V1','Datum':'2026-06-01'},
            {'OtpremnicaID':'OTP-2','StanicaID':'ST-1','BrojOtpremnice':'102','Stornirano':'Ne','VozacID':'V2','Datum':'2026-06-01'},
            {'OtpremnicaID':'OTP-3','StanicaID':'ST-1','BrojOtpremnice':'103','Stornirano':'Ne','VozacID':'V3','Datum':'2026-06-01'}]
    srcP=[{**r,'BrojZbirne':'Z1' if i==1 else 'Z2'} for i,r in enumerate(canonP,1)]
    canonZ=[{'ZbirnaID':'ZB-1','BrojZbirne':'Z1','VozacID':'V1','Datum':'2026-06-01','Stornirano':'Ne'},
            {'ZbirnaID':'ZB-2','BrojZbirne':'Z2','VozacID':'V2','Datum':'2026-06-01','Stornirano':'Ne'},
            {'ZbirnaID':'ZB-3','BrojZbirne':'Z2','VozacID':'V3','Datum':'2026-06-01','Stornirano':'Ne'}]
    for t,c,s in [('tblOtkup',canonO,srcO),('tblOtpremnica',canonP,srcP),('tblZbirna',canonZ,canonZ)]:
        save(base/(t+'.jsonl'),c);save(base/('SOURCE__'+t+'.jsonl'),s)

with tempfile.TemporaryDirectory() as d:
    root=Path(d);stage=root/'original';fixture(stage)
    out=root/'output'
    assert m.main(['--stage',str(stage),'--out',str(out),'--stage-copy'])==0 # one historically unlinked, documented
    report=json.loads((out/'mig11_build_summary.json').read_text())
    assert report['new_link_rows']=={'tblOtpremnicaIzvori':3,'tblZbirnaIzvori':3},report
    assert report['unresolved']=={}
    assert report['historically_unlinked']=={'OPI_NO_HISTORICAL_OTP_REFERENCE':1}
    assert report['historical_warnings']=={'OPI_CANCELLED_OTK_ACTIVE_OTP':1,'OPI_CANCELLED_SOURCE':1}
    assert report['snapshot']=='stage_MIG11_ready'
    assert len(m.stage_rows(out/'stage_MIG11_ready'/'tblOtpremnicaIzvori.jsonl'))==3
    assert len(m.stage_rows(out/'stage_MIG11_ready'/'tblZbirnaIzvori.jsonl'))==3
    for t in ('tblOtkup','tblOtpremnica','tblZbirna'):
        assert m.hashfile(stage/(t+'.jsonl'))==m.hashfile(out/'stage_MIG11_ready'/(t+'.jsonl'))
    try:
        with redirect_stderr(StringIO()):
            m.main(['--stage',str(stage),'--out',str(out)])
        raise AssertionError('must reject existing output')
    except SystemExit as e:
        assert e.code==2
    # exactly same names -> deterministic row values
    opi,zbi,_,_,_,_=m.build(m.load(stage)[0]); opi2,zbi2,_,_,_,_=m.build(m.load(stage)[0]);assert opi==opi2 and zbi==zbi2
    # existing link must be preserved not duplicated
    save(stage/'tblOtpremnicaIzvori.jsonl',opi[:1]);save(stage/'tblZbirnaIzvori.jsonl',zbi[:1])
    opi3,zbi3,issues3,_,_,unlinked3=m.build(m.load(stage)[0]); assert len(opi3)==3 and len(zbi3)==3 and len(issues3)==0 and unlinked3['OPI_NO_HISTORICAL_OTP_REFERENCE']==1
    # source target broken -> issue not fabricated
    rows=m.stage_rows(stage/'SOURCE__tblOtkup.jsonl');rows[0]['OtpremnicaID']='OTP-missing';save(stage/'SOURCE__tblOtkup.jsonl',rows)
    _,_,errors,_,_,_=m.build(m.load(stage)[0]);assert any(e['reason']=='DANGLING_OTP_REFERENCE' for e in errors)
    # contradictory Zbirna references must not become a link.
    rows=m.stage_rows(stage/'SOURCE__tblOtkup.jsonl');rows[0]['OtpremnicaID']='OTP-1';rows[0]['BrojZbirne']='WRONG';save(stage/'SOURCE__tblOtkup.jsonl',rows)
    _,_,errors,_,_,_=m.build(m.load(stage)[0]); assert any(e['reason']=='CONFLICT_BROJ_ZBIRNE' for e in errors)
    # source mismatches canon explicitly abort
    rows=m.stage_rows(stage/'SOURCE__tblOtkup.jsonl');rows[0]['StanicaID']='OTHER';save(stage/'SOURCE__tblOtkup.jsonl',rows)
    try: m.build(m.load(stage)[0]);raise AssertionError('must have failed')
    except ValueError as e: assert 'konflikt' in str(e)
    # ambiguous reused business number -> never guess
    assert m.lookup_zbr('Z2',{'BrojZbirne':'Z2'},m.stage_rows(stage/'tblZbirna.jsonl'))==(None,'AMBIGUOUS_NUMBER')
    # conflicting date and driver -> ambiguous
    target,how=m.lookup_zbr('Z2',{'VozacID':'V2','Datum':'2027-10-15'},m.stage_rows(stage/'tblZbirna.jsonl'))
    assert how=='UNIQUE_NUMBER_DRIVER',how
    target,how=m.lookup_zbr('Z2',{'VozacID':'wrong','Datum':'2026-06-01'},m.stage_rows(stage/'tblZbirna.jsonl'))
    assert target is None,how

with tempfile.TemporaryDirectory() as d:
    root=Path(d); stage=root/'stage'; fixture(stage)
    rows=m.stage_rows(stage/'SOURCE__tblOtkup.jsonl')
    rows[0]['OtpremnicaID']='OTP-MISSING'
    save(stage/'SOURCE__tblOtkup.jsonl',rows)
    assert m.main(['--stage',str(stage),'--out',str(root/'result'),'--stage-copy'])==0
    report=json.loads((root/'result'/'mig11_build_summary.json').read_text())
    assert report['status']=='PARTIAL_LINKS_NEED_REVIEW'
    assert report['snapshot']=='BLOCKED_BY_RELATION_ERRORS'
    assert not (root/'result'/'stage_MIG11_ready').exists()
    assert report['unresolved']=={'OPI:DANGLING_OTP_REFERENCE':1}

print('MIG11 BUILD LINK synthetic smoke PASS (happy path, unchanged input, blocked invalid FK, ambiguous label, deterministic output)')
