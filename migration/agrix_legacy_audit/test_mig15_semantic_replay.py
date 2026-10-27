import hashlib
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import agrix_mig15_semantic_replay as m


class ReplayTest(unittest.TestCase):
    def test_legacy_balance_crossing_without_input_changes(self):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td)/'project';out=Path(td)/'result'
            base=[
                {'physical_case_id':'A','type':'GAJBA','from_account_type':'Kooperant','from_account_id':'K1',
                 'to_account_type':'Stanica','to_account_id':'S1','quantity':2},
                {'physical_case_id':'B','type':'GAJBA','from_account_type':'Stanica','from_account_id':'S1',
                 'to_account_type':'Kooperant','to_account_id':'K1','quantity':3},
                {'physical_case_id':'C','type':'GAJBA','from_account_type':'Stanica','from_account_id':'S1',
                 'to_account_type':'Vozac','to_account_id':'V1','quantity':1},
            ]
            def ledger(case,date,kind,fr,fi,to,ti,qty):
                return {'physical_case_id':case,'status':'NOT_APPROVED','source_row_ordinals':[1 if case=='A' else 2],
                        'reserved_AmbID_NOT_ALLOCATED':'AMB-'+case,
                        'target_column_values_WITHOUT_AMBID':{
                            'Datum':date,'TipAmbalaze':'GAJBA','Kolicina':qty,'DokumentID':'D'+case,'DokumentTIP':'Revers',
                            'OdNalogTip':fr,'OdNalogID':fi,'NaNalogTip':to,'NaNalogID':ti,'VrstaKretanja':kind,'StornoOd':''}}
            ledger_rows=[ledger('A','2026-08-01','POVRAT_PRAZNE','Kooperant','K1','Stanica','S1',2),
                        ledger('B','2026-08-02','IZDATA_PRAZNA','Stanica','S1','Kooperant','K1',3)]
            reserved=[{'physical_case_id':x['physical_case_id'],
                       'candidate_AmbID_NOT_ALLOCATED':x['reserved_AmbID_NOT_ALLOCATED'],
                       'movement_kind_proposal':x['target_column_values_WITHOUT_AMBID']['VrstaKretanja']} for x in ledger_rows]
            holds=[{'physical_case_id':'C'}]
            hdr=[{'proposed_target_header':{'AmbDokID':'ADK-1'}}]
            raw=[{'AmbID':'legacy-1'}]
            files={
            'MigrationAmb10BusinessIntent_v2/AMB10_BUSINESS_INTENT_ALL_CASES_PRIVATE.jsonl':base,
            'MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_TARGET_ROWS_NOT_APPROVED_PRIVATE.jsonl':ledger_rows,
            'MigrationAmb10TargetAdapter_v2/AMB10_OFFLINE_TARGET_HEADERS_NOT_APPROVED_PRIVATE.jsonl':hdr,
            'MigrationAmb10ImportReadiness_v2/AMB10_CANDIDATE_ID_AND_SCOPED_GATES_PRIVATE.jsonl':reserved,
            'MigrationAmb10ImportReadiness_v2/AMB10_UNRESOLVED_DOCUMENT_HOLDS_PRIVATE.jsonl':holds,
            'MigrationStage_v2/UNCONVERTED__tblAmbalaza.jsonl':raw}
            tmp_expected={}
            for rel,rows in files.items():
                path=root/rel;path.parent.mkdir(parents=True,exist_ok=True)
                path.write_text(''.join(json.dumps(x)+'\n' for x in rows),encoding='utf-8')
                tmp_expected[rel]=(hashlib.sha256(path.read_bytes()).hexdigest(),len(rows))
            before={path:path.read_bytes() for path in [root/rel for rel in files]}
            with patch.dict(m.EXPECTED,tmp_expected,clear=True):
                rc=m.run(root,out)
                self.assertEqual(rc,0)
                result=json.loads((out/'MIG15_semantic_replay_summary.json').read_text())
                self.assertTrue(result['balance_conservation_reconstructed_plus_holds_matches_all_physical'])
                self.assertEqual(result['historical_crossings']['ISSUANCE_CROSSES_ZERO'],1)
                self.assertEqual(result['historical_crossings']['RETURN_ENDS_WITH_NEGATIVE_BALANCE'],1)
                self.assertEqual(result['historical_crossing_cases_private'],2)
                private=(out/'MIG15_crossings_PRIVATE.jsonl').read_text()
                self.assertIn('potential_legacy_due_part',private)
                self.assertFalse(result['source_and_target_workbooks_modified'])
                for path, content in before.items():self.assertEqual(path.read_bytes(),content)
                (root/'MigrationAmb10BusinessIntent_v2/AMB10_BUSINESS_INTENT_ALL_CASES_PRIVATE.jsonl').write_text('tampered')
                with self.assertRaises(ValueError):m.run(root,out/'invalid')

if __name__=='__main__':unittest.main()
