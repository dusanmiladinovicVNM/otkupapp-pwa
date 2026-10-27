import importlib.util
import hashlib
import json
import tempfile
from pathlib import Path

spec=importlib.util.spec_from_file_location('migration',str(Path(__file__).with_name('agrix_mig12_archive_mig13_finance.py')))
m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)

def make_jsonl(path,rows):
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(''.join(json.dumps(r,ensure_ascii=False)+'\n' for r in rows),encoding='utf-8')

with tempfile.TemporaryDirectory() as td:
    root=Path(td); stage=root/'stage'; stage.mkdir(); out=root/'out'
    a=[{'AmbID':'A'+str(i),'Stornirano':'Da' if i<335 else ''} for i in range(7815)]
    st=[{'CorrectionID':'C'+str(i)} for i in range(2)]
    n=[{'NovacID':f'NOV-{i}', 'Datum':'2026-08-01', 'Tip':'UPLATA',
        'Partner':'Partner test','EntitetTip':'Kupac','Uplata':'123.45','Isplata':0}
       for i in range(39)]
    for name,rows in [('UNCONVERTED__tblAmbalaza.jsonl',a),('UNCONVERTED__tblStornoVeze.jsonl',st),('SECONDARY__tblNovac.jsonl',n)]:
        p=stage/name;make_jsonl(p,rows);m.EXPECTED[name]=(m.sha(p),len(rows))
    m.EXPECTED_AMB_SOURCE='test-source-fingerprint'
    for relative in m.AMB_REPORTS:
        p=root/relative;p.parent.mkdir(parents=True,exist_ok=True)
        p.write_text(json.dumps({'source_sha256':'test-source-fingerprint','status':'READ_ONLY',
           'import_approved':False,'row_approval_count':0}),encoding='utf-8')
    r=m.execute(stage,out,root,None)
    assert r['archive']['status']=='PASS'
    assert r['finance']['rows']==39
    assert r['finance']['ready_for_candidate']
    assert (out/'MIG13_tblNovac_CANDIDATE_NOT_IMPORTED.jsonl').exists()
    assert m.sha(out/'MIG13_SECONDARY_tblNovac_SOURCE.jsonl')==m.EXPECTED['SECONDARY__tblNovac.jsonl'][0]
    assert r['archive']['historical_reports_preserved']==10
    try:m.execute(stage,out,root,None)
    except ValueError as e:assert 'Output already exists' in str(e)
    else:raise AssertionError('rerun must fail closed')
    n[0]['Uplata']=-1
    s,detail=m.preflight(n,stage)
    assert not s['ready_for_candidate'] and s['issue_counts']['NEGATIVE_AMOUNT']==1
    n[0]['Uplata']=5
    n[0]['OtkupID']='fake'
    s,detail=m.preflight(n,stage)
    assert not s['ready_for_candidate'] and any('FK_TABLE_NOT_AVAILABLE_OTKUPID' in d['issue_codes'] for d in detail)
    n[0].pop('OtkupID')
    n[0]['Partner']=''
    s,detail=m.preflight(n,stage)
    assert s['issue_counts']['MISSING_PARTNER']==1
    n[0]['Partner']='Partner test'
    n[0]['UnexpectedNonEmpty']='critical'
    s,detail=m.preflight(n,stage)
    assert s['issue_counts']['UNRECOGNIZED_NONEMPTY_COLUMN']==1
    assert not s['ready_for_candidate']
    print('TESTS PASS: archive SHA, 7815+2 source rows, 39 finance, 10 reports, no overwrite, negative, missing FK, partner, unknown field')
