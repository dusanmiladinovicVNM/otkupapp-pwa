import json, tempfile, sys
from pathlib import Path
from decimal import Decimal
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).parent))
import agrix_mig173_trace_missing_payment as m

assert m.legacy_payment_value('DA') == (None, 'non_numeric')
assert m.legacy_payment_value('NE') == (None, 'non_numeric')
assert m.legacy_payment_value(True) == (None, 'non_numeric')
assert m.legacy_payment_value(None) == (Decimal(0), 'blank')
assert m.legacy_payment_value('100000') == (Decimal('100000'), 'numeric')
assert m.legacy_payment_value('100.000,00') == (Decimal('100000.00'), 'numeric_eu')
assert m.legacy_payment_value('100 000,00') == (Decimal('100000.00'), 'numeric_eu')
assert m.legacy_payment_value('100000,50') == (Decimal('100000.50'), 'numeric')
rows = [
  {'KooperantID':'K-1', 'Stornirano':'', 'Isplaceno':'DA', 'OtkupID':'OTK-1'},
  {'KooperantID':'K-1', 'Stornirano':'', 'Isplaceno':'100.000,00', 'OtkupID':'OTK-2'},
  {'KooperantID':'K-1', 'Stornirano':'Da', 'Isplaceno':'300000', 'OtkupID':'OTK-3'},
  {'KooperantID':'K-2', 'Stornirano':'', 'Isplaceno':'NE', 'OtkupID':'OTK-4'},
]
paid, unknown, counts = m.candidate_old_otkup(rows, 'K-1')
assert len(paid)==1 and paid[0][1] == 100000
assert len(unknown)==1 and counts['non_numeric']==1
assert len(rows)==4

with tempfile.TemporaryDirectory() as temp:
    base = Path(temp)
    stage=base/'MigrationStage_v2';stage.mkdir()
    with (stage/'SOURCE__tblOtkup.jsonl').open('w',encoding='utf-8') as f:
        for row in rows:f.write(json.dumps(row)+'\n')
    novac=[]
    for i in range(39):
        r={'NovacID':f'NOV-{i:03}', 'KooperantID':'K-1' if i==0 else 'K-2',
           'Isplata':'100000' if i==0 else '0', 'OMID':'ST-1', 'Datum':'2026-06-15'}
        novac.append(r)
    with (stage/'SECONDARY__tblNovac.jsonl').open('w',encoding='utf-8') as f:
        for row in novac:f.write(json.dumps(row)+'\n')
    workbook=base/'QA.xlsm';workbook.write_bytes(b'dummy-nonexcel-for-mocked-reader')
    fake={'tblNovac':novac,'tblOtkup':[],
          'tblKooperanti':[{'Ime':'Dragan','Prezime':'Milosavljevic','KooperantID':'K-1','StanicaID':'ST-1'}]}
    out=base/'result'
    with patch.object(m,'read_book',return_value=(fake,{})):
        s=m.run(base,workbook,out,'ST-1')
    assert s['status']=='READ_ONLY_TRACE_WITH_UNRESOLVED_SOURCE_FIELD'
    assert s['legacy_otkup_non_numeric_isplaceno_count']==1
    assert s['legacy_otkup_numeric_candidates_total_100000']
    assert s['canonical_novac_subject_positive_rows']==1
    assert not s['legacy_otkup_payment_total_proven']
    assert workbook.read_bytes()==b'dummy-nonexcel-for-mocked-reader'
    assert (out/'MIG173_missing_payment_PRIVATE.json').exists()
    assert (out/'MIG173_summary.json').exists()
    try:
      with patch.object(m,'read_book',return_value=(fake,{})):
          m.run(base,workbook,out,'ST-1')
      raise AssertionError('Must refuse output overwrite')
    except ValueError as e: assert 'already exists' in str(e)
print('MIG17.3 v1.1 tests PASS: 11 targeted + one mocked end-to-end + overwrite guard')
