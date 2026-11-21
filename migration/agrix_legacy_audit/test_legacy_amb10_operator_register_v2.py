import csv
import json
from pathlib import Path
import tempfile
import unittest
from legacy_audit import AuditError
import legacy_amb10_operator_register_v2 as o


def write_csv(path, rows):
    with path.open('w',encoding='utf-8-sig',newline='') as fd:
        w=csv.DictWriter(fd,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)


def fixture(folder,pairs=43,single=12,negative=35):
    folder.mkdir()
    qs=[]
    for i in range(1,pairs*2+single+1):
        peer=i+1 if i<=2*pairs and i%2 else (i-1 if i<=2*pairs else '')
        qs.append({'source_row_ordinal':str(i),'physical_peer_row_candidate':str(peer),
                   'quarantine_case':'DOCUMENT_ID_NOT_FOUND' if i<=2*pairs else 'DOCUMENT_TYPE_NOT_MODELLED',
                   'document_tip':'OM-Izlaz-Koop','entitet_id':'ID-PRIVATE-'+str(i),
                   'document_id':'PRIVATE-DOC', 'operator_resolution':'UNREVIEWED'})
    ns=[{'account_type':'Kooperant','account_id':f'PERSON-{i}', 'tip_ambalaze':'12/1'} for i in range(negative)]
    write_csv(folder/'amb10_quarantine_operator_PRIVATE.csv',qs)
    write_csv(folder/'amb10_negative_accounts_operator_PRIVATE.csv',ns)
    (folder/'agrix_amb10_quarantine_summary.json').write_text(json.dumps({
        'format':'MIG-04.4','status':'PRIVATE_TRIAGE_ONLY_NO_AMB10_APPROVAL_OR_IMPORT',
        'quarantined_active_rows':len(qs),'negative_legacy_account_type_combinations':len(ns),
        'exact_reciprocal_old_legs_candidate_pairs':pairs,'old_legs_without_unique_reciprocal_evidence':single,
        'source_sha256':'TEST-SHA','schema_sha256':'TEST-SCHEMA'}))


class MIG045(unittest.TestCase):
    def test_partition_98_to_55_and_neg_35(self):
        with tempfile.TemporaryDirectory() as t:
            folder=Path(t)/'dossier';fixture(folder)
            src,cases,neg,dig=o.dossier_load(folder)
            self.assertEqual(len(cases),55);self.assertEqual(len(neg),35)
            self.assertEqual(sum(z['physical_case']=='RECIPROCAL_TWO_LEGS' for z in cases),43)
            self.assertEqual(sum(z['physical_case']=='SINGLE_UNPAIRED_LEG' for z in cases),12)
            self.assertNotIn('PERSON-',json.dumps(o.safe_summary(src,cases,neg,{}, {}, 'prepare',[])))

    def test_preparation_and_check_no_authorization(self):
        with tempfile.TemporaryDirectory() as t:
            d=Path(t)/'dossier';fixture(d)
            p=Path(t)/'decision';c=Path(t)/'check'
            self.assertEqual(o.cli(['--mode','prepare','--dossier',str(d),'--out',str(p)]),0)
            self.assertEqual(o.cli(['--mode','check','--dossier',str(d),'--decisions',str(p),'--out',str(c)]),0)
            result=json.loads((c/'agrix_amb10_operator_summary.json').read_text())
            self.assertEqual(result['operator_physical_cases'],55)
            self.assertEqual(result['negative_balance_account_cases'],35)
            self.assertEqual(result['movement_review_status_counts'],{'PENDING':55})
            self.assertFalse(result['ready_for_import'])
            self.assertEqual(result['validation_issue_counts'],{})
            self.assertNotIn('PRIVATE-DOC',json.dumps(result))

    def test_documents_cannot_be_approved_without_independent_evidence(self):
        with tempfile.TemporaryDirectory() as t:
            d=Path(t)/'dossier';fixture(d,pairs=1,single=1,negative=1)
            p=Path(t)/'decision';o.cli(['--mode','prepare','--dossier',str(d),'--out',str(p)])
            filepath=p/'amb10_movement_decisions_PRIVATE.csv'
            rows=o.read_csv(filepath,o.Q_FIELDS)
            rows[0]['operator_resolution']='DOCUMENTED_FOR_REVIEW'
            write_csv(filepath,rows)
            result,*_=o.build_results(d,'check',p)
            self.assertIn('MOVEMENT_INDEPENDENT_PROOF_MISSING',result['validation_issue_counts'])
            rows[0].update(direction_verified='YES',movement_kind='PRENOS_INTERNO',
                           packaging_ownership='OWN_FIRM',source_document_basis='INDEPENDENT_SOURCE_DOCUMENT',
                           independent_evidence_reference='PAPER-15',operator_initials='OP')
            write_csv(filepath,rows)
            result,*_=o.build_results(d,'check',p)
            self.assertEqual(result['validation_issue_counts'],{})
            self.assertFalse(result['ready_for_import'])

    def test_fail_closed_missing_mutated_or_duplicate(self):
        with tempfile.TemporaryDirectory() as t:
            d=Path(t)/'dossier';fixture(d,pairs=1,single=1,negative=1)
            p=Path(t)/'decision';o.cli(['--mode','prepare','--dossier',str(d),'--out',str(p)])
            filepath=p/'amb10_movement_decisions_PRIVATE.csv'
            rows=o.read_csv(filepath,o.Q_FIELDS)
            rows[0]['source_row_ordinals']='100+200'
            write_csv(filepath,rows)
            with self.assertRaisesRegex(AuditError,'Read-only'):o.build_results(d,'check',p)
            rows[0]['source_row_ordinals']='1+2'
            rows[1]['case_id']=rows[0]['case_id']
            write_csv(filepath,rows)
            with self.assertRaisesRegex(AuditError,'duplicate'):o.build_results(d,'check',p)

    def test_source_dossier_change_refused(self):
        with tempfile.TemporaryDirectory() as t:
            d=Path(t)/'dossier';fixture(d,pairs=1,single=0,negative=1)
            p=Path(t)/'decision';o.cli(['--mode','prepare','--dossier',str(d),'--out',str(p)])
            with (d/'amb10_quarantine_operator_PRIVATE.csv').open('a') as fd:fd.write('\n')
            with self.assertRaisesRegex(AuditError,'changed'):o.build_results(d,'check',p)

    def test_nonreciprocal_pair_refused(self):
        with tempfile.TemporaryDirectory() as t:
            d=Path(t)/'dossier';fixture(d,pairs=1,single=0,negative=1)
            f=d/'amb10_quarantine_operator_PRIVATE.csv'
            rows=o.read_csv(f,('source_row_ordinal',))
            rows[1]['physical_peer_row_candidate']='3';write_csv(f,rows)
            with self.assertRaisesRegex(AuditError,'Nonreciprocal'):o.dossier_load(d)

    def test_opening_evidence_required(self):
        with tempfile.TemporaryDirectory() as t:
            d=Path(t)/'dossier';fixture(d,pairs=1,single=0,negative=1)
            p=Path(t)/'decision';o.cli(['--mode','prepare','--dossier',str(d),'--out',str(p)])
            f=p/'amb10_negative_decisions_PRIVATE.csv';rows=o.read_csv(f,o.N_FIELDS)
            rows[0]['operator_resolution']='EVIDENCE_PROVIDED';write_csv(f,rows)
            report,*_=o.build_results(d,'check',p)
            self.assertIn('NEGATIVE_ACCOUNT_INDEPENDENT_PROOF_MISSING',report['validation_issue_counts'])

    def test_cannot_overwrite_and_output_overlap(self):
        with tempfile.TemporaryDirectory() as t:
            d=Path(t)/'dossier';fixture(d,pairs=1,single=0,negative=1)
            p=Path(t)/'decisions'
            o.cli(['--mode','prepare','--dossier',str(d),'--out',str(p)])
            with self.assertRaisesRegex(AuditError,'exists'):
                o.cli(['--mode','prepare','--dossier',str(d),'--out',str(p)])
            with self.assertRaisesRegex(AuditError,'overlap'):
                o.cli(['--mode','prepare','--dossier',str(d),'--out',str(d/'new')])

if __name__=='__main__':unittest.main()
