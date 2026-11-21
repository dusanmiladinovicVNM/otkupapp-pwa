import json,tempfile,hashlib,copy
from pathlib import Path
import agrix_mig16_daily_close as m

def jsonl(path,rows):
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(''.join(json.dumps(r)+'\n' for r in rows),encoding='utf-8')
    return (hashlib.sha256(path.read_bytes()).hexdigest(),len(rows))

with tempfile.TemporaryDirectory() as t:
    root=Path(t)/'project';out=root/'MigrationAmb10_MIG16'
    spec={}
    bas=[];led=[];res=[];hold=[]
    for n in range(4675):
        case=f'CASE{n:04d}';typ='Gajbica'
        frm=('Kooperant','K1') if n%2 else ('Stanica','S1')
        to=('Stanica','S1') if n%2 else ('Kooperant','K1')
        bas.append(dict(physical_case_id=case,type=typ,quantity=1,from_account_type=frm[0],from_account_id=frm[1],to_account_type=to[0],to_account_id=to[1]))
        if n<4669:
            d=dict(Datum='2026-08-01',TipAmbalaze=typ,Kolicina=1,OdNalogTip=frm[0],OdNalogID=frm[1],NaNalogTip=to[0],NaNalogID=to[1],VrstaKretanja=('POVRAT_PRAZNE' if n%2 else 'IZDATA_PRAZNA'))
            led.append(dict(physical_case_id=case,reserved_AmbID_NOT_ALLOCATED=f'MIG-{n}',target_column_values_WITHOUT_AMBID=d))
            res.append(dict(physical_case_id=case,candidate_AmbID_NOT_ALLOCATED=f'MIG-{n}',movement_kind_proposal=d['VrstaKretanja']))
        else:hold.append(dict(physical_case_id=case))
    m.EXPECTED={
     'MigrationAmb10BusinessIntent_v2/AMB10_BUSINESS_INTENT_ALL_CASES_PRIVATE.jsonl':jsonl(root/'MigrationAmb10BusinessIntent_v2/AMB10_BUSINESS_INTENT_ALL_CASES_PRIVATE.jsonl',bas),
     'MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_TARGET_ROWS_NOT_APPROVED_PRIVATE.jsonl':jsonl(root/'MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_TARGET_ROWS_NOT_APPROVED_PRIVATE.jsonl',led),
     'MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_TARGET_HEADERS_NOT_APPROVED_PRIVATE.jsonl':jsonl(root/'MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_TARGET_HEADERS_NOT_APPROVED_PRIVATE.jsonl',[dict(proposed_target_header={'AmbDokID':f'DK-{i}'}) for i in range(48)]),
     'MigrationAmb10ImportReadiness_v2/AMB10_CANDIDATE_ID_AND_SCOPED_GATES_PRIVATE.jsonl':jsonl(root/'MigrationAmb10ImportReadiness_v2/AMB10_CANDIDATE_ID_AND_SCOPED_GATES_PRIVATE.jsonl',res),
     'MigrationAmb10ImportReadiness_v2/AMB10_UNRESOLVED_DOCUMENT_HOLDS_PRIVATE.jsonl':jsonl(root/'MigrationAmb10ImportReadiness_v2/AMB10_UNRESOLVED_DOCUMENT_HOLDS_PRIVATE.jsonl',hold),
     'MigrationStage_v2/UNCONVERTED__tblAmbalaza.jsonl':jsonl(root/'MigrationStage_v2/UNCONVERTED__tblAmbalaza.jsonl',[{} for i in range(7815)]),
    }
    m.EXTRA={
     'MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_EXCEPTIONS_NOT_APPROVED_PRIVATE.jsonl':jsonl(root/'MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_EXCEPTIONS_NOT_APPROVED_PRIVATE.jsonl',[{} for i in range(8)]),
     'MigrationAmb10ReversDocumentPlan_v2/AMB10_REVERS_EVENT_LINK_DRAFT_PRIVATE.jsonl':jsonl(root/'MigrationAmb10ReversDocumentPlan_v2/AMB10_REVERS_EVENT_LINK_DRAFT_PRIVATE.jsonl',[{} for i in range(48)]),
    }
    summary=m.run(root,out)
    assert summary['candidate_physical_events_checked']==4669
    assert summary['physical_cases_not_yet_in_canonical_candidate']==6
    assert summary['canonical_import_approved'] is False
    assert (out/'MIG16_daily_close_PRIVATE.csv').exists()
    assert (out/'MIG16_summary.json').exists()
    assert not list(root.rglob('*.xlsm'))
    try:m.run(root,out)
    except ValueError as e:assert 'already exists' in str(e)
    else:raise AssertionError('Output overwrite must block')
    print('MIG16 PIPELINE SIMULATION PASS')
