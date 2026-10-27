import json
import tempfile
from pathlib import Path
import importlib.util
from contextlib import redirect_stdout, redirect_stderr
from io import StringIO

p=Path(__file__).resolve().with_name('agrix_mig11_build_links.py')
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
    opi,zbi,_,_,_,_,_=m.build(m.load(stage)[0]); opi2,zbi2,_,_,_,_,_=m.build(m.load(stage)[0]);assert opi==opi2 and zbi==zbi2
    # existing link must be preserved not duplicated
    save(stage/'tblOtpremnicaIzvori.jsonl',opi[:1]);save(stage/'tblZbirnaIzvori.jsonl',zbi[:1])
    opi3,zbi3,issues3,_,_,unlinked3,_=m.build(m.load(stage)[0]); assert len(opi3)==3 and len(zbi3)==3 and len(issues3)==0 and unlinked3['OPI_NO_HISTORICAL_OTP_REFERENCE']==1
    # source target broken -> issue not fabricated
    rows=m.stage_rows(stage/'SOURCE__tblOtkup.jsonl');rows[0]['OtpremnicaID']='OTP-missing';save(stage/'SOURCE__tblOtkup.jsonl',rows)
    _,_,errors,_,_,_,_=m.build(m.load(stage)[0]);assert any(e['reason']=='DANGLING_OTP_REFERENCE' for e in errors)
    # An explicit OtpremnicaID is preserved despite a contradictory *display* BrojZbirne.
    rows=m.stage_rows(stage/'SOURCE__tblOtkup.jsonl');rows[0]['OtpremnicaID']='OTP-1';rows[0]['BrojZbirne']='WRONG';save(stage/'SOURCE__tblOtkup.jsonl',rows)
    resolved_opi,_,errors,_,risk,_,mismatch=m.build(m.load(stage)[0])
    assert len(resolved_opi)==3 and not errors
    assert risk['OPI_LEGACY_BROJ_ZBIRNE_MISMATCH']==1 and len(mismatch)==1
    assert mismatch[0]['OtkupID']=='OTK-1' and mismatch[0]['OtpremnicaID']=='OTP-1'
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

with tempfile.TemporaryDirectory() as d:
    root=Path(d);stage=root/'original';fixture(stage)
    orig_hash={p.name:m.hashfile(p) for p in stage.iterdir()}
    rows=m.stage_rows(stage/'SOURCE__tblOtkup.jsonl')
    rows[0]['BrojZbirne']='WRONG'
    save(stage/'SOURCE__tblOtkup.jsonl',rows)
    orig_hash['SOURCE__tblOtkup.jsonl']=m.hashfile(stage/'SOURCE__tblOtkup.jsonl')
    out=root/'result'
    assert m.main(['--stage',str(stage),'--out',str(out),'--stage-copy'])==0
    report=json.loads((out/'mig11_build_summary.json').read_text())
    assert report['format']=='AGRIx-MIG-11-LINKS-BUILD-v1.1'
    assert report['status']=='READY_WITH_HISTORICAL_NOTES'
    assert report['new_link_rows']=={'tblOtpremnicaIzvori':3,'tblZbirnaIzvori':3}
    assert report['unresolved']=={}
    assert report['historical_number_mismatches_preserved_by_explicit_id']==1
    assert report['manual_historical_review_required'] is True
    assert report['snapshot']=='stage_MIG11_ready'
    assert m.stage_rows(out/'stage_MIG11_ready'/'tblOtpremnicaIzvori.jsonl')[0]['OtpremnicaID']=='OTP-1'
    assert len((out/'mig11_number_mismatches_PRIVATE_LOCAL.csv').read_text(encoding='utf-8-sig').splitlines())==2
    assert {p.name:m.hashfile(p) for p in stage.iterdir()}==orig_hash
    # A canonical link conflicting with explicit historical direct ID is still a blocker.
    save(stage/'tblOtpremnicaIzvori.jsonl',[{'OtpremnicaIzvorID':'OPI-existing','OtpremnicaID':'OTP-2','OtkupID':'OTK-1'}])
    out2=root/'result_blocked'
    assert m.main(['--stage',str(stage),'--out',str(out2),'--stage-copy'])==0
    assert (json.loads((out2/'mig11_build_summary.json').read_text())['snapshot']=='BLOCKED_BY_RELATION_ERRORS')
    assert 'OPI:CONFLICT_WITH_EXISTING_LINK' in json.loads((out2/'mig11_build_summary.json').read_text())['unresolved']

print('MIG11 v1.1 PASS: direct-ID/display-number discrepancy preserved, private audit generated, stage-copy gated, conflicting canonical links blocked')

# Canonical data in a nested folder: newly generated links must go there,
# not be silently dropped at staging root where MIG-09 might not scan them.
with tempfile.TemporaryDirectory() as d:
    root=Path(d); stage=root/'stage';stage.mkdir();(stage/'canonical').mkdir()
    fixture_stage=root/'flat';fixture(fixture_stage)
    for p in fixture_stage.iterdir():
        (stage/'canonical'/p.name).write_bytes(p.read_bytes())
    out=root/'out'
    assert m.main(['--stage',str(stage),'--out',str(out),'--stage-copy'])==0
    assert (out/'stage_MIG11_ready'/'canonical'/'tblOtpremnicaIzvori.jsonl').exists()
    assert (out/'stage_MIG11_ready'/'canonical'/'tblZbirnaIzvori.jsonl').exists()
    assert not (out/'stage_MIG11_ready'/'tblOtpremnicaIzvori.jsonl').exists()
print('MIG11 v1.1 PASS: nested canonical staging layout preserved')
