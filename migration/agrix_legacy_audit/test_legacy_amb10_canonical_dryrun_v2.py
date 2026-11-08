import io
import json
import tempfile
import unittest
from pathlib import Path
from contextlib import redirect_stdout
from unittest.mock import patch

from legacy_audit import AuditError
from legacy_amb10_canonical_dryrun_v2 import canonical_preview, _ledger_delta, cli, VERSION
import legacy_seven_hold_rescue_v2 as prior
import test_legacy_seven_hold_rescue_v2 as previous_tests


def example(kind='AMBALAZA_UZ_ROBU',status='DIRECTION_KIND_PROPOSED_ONLY'):
    return {'physical_case_id':'PHY-00001','legacy_DokumentTIP':'Otkup',
            'legacy_DokumentID':'OTK-1','legacy_Datum':'45500',
            'TipAmbalaze':'Gajba','Kolicina':'5.25',
            'OdNalogTip':'Kooperant','OdNalogID':'K1',
            'NaNalogTip':'Stanica','NaNalogID':'S1',
            'VrstaKretanja_PROPOSAL':kind,'evidence_status':status,
            'canonical_AMB10_import_authorized':'NO'}


class UnitTests(unittest.TestCase):
    def setUp(self):
        self.parents={'SOURCE__tblOtkup':{'OTK-1':(1,{'Stornirano':''})},
                      'SOURCE__tblOtpremnica':{},'tblPrijemnica':{}}
    def test_target_fields_are_proposed_not_approved(self):
        row,reasons=canonical_preview(example(),self.parents)
        self.assertFalse(reasons)
        self.assertEqual(row['DokumentTIP'],'Otkup')
        self.assertEqual(row['Kolicina'],'5.25')
        self.assertEqual(row['Datum'],'2024-07-27')
        self.assertNotIn('AmbID',row)
        self.assertNotIn('StornoOd',row)
    def test_block_old_revers(self):
        e=example();e['legacy_DokumentTIP']='OM-Izlaz-Koop';e['legacy_DokumentID']='DOC-1';e['VrstaKretanja_PROPOSAL']='UNASSIGNED';e['evidence_status']='HOLD_OWNERSHIP'
        v,why=canonical_preview(e,self.parents)
        self.assertIsNone(v)
        self.assertIn('CANONICAL_SOURCE_DOCUMENT_IDENTITY_UNRESOLVED',why)
    def test_missing_parent(self):
        e=example();e['legacy_DokumentID']='OTK-9'
        row,reasons=canonical_preview(e,self.parents)
        self.assertIsNone(row)
        self.assertIn('CANONICAL_SOURCE_PARENT_NOT_FOUND',reasons)
    def test_storno_parent(self):
        self.parents['SOURCE__tblOtkup']['OTK-1']=(1,{'Stornirano':'Da'})
        row,reasons=canonical_preview(example(),self.parents)
        self.assertIsNone(row)
        self.assertIn('CANONICAL_SOURCE_PARENT_STORNIRANO',reasons)
    def test_reject_import_flag(self):
        e=example();e['canonical_AMB10_import_authorized']='YES'
        with self.assertRaises(AuditError):canonical_preview(e,self.parents)
    def test_no_forged_date(self):
        e=example();e['legacy_Datum']='07/08/2024'
        row,reasons=canonical_preview(e,self.parents)
        self.assertIsNone(row)
        self.assertIn('DATE_UNPARSEABLE',reasons)
    def test_movement_matrix(self):
        e=example(kind='POVRAT_PRAZNE');e['OdNalogTip']='Stanica';e['NaNalogTip']='Kooperant'
        row,reasons=canonical_preview(e,self.parents)
        self.assertIsNone(row)
        self.assertIn('KIND_ACCOUNT_MATRIX_CONFLICT',reasons)
    def test_allow_negative_by_construction(self):
        from collections import defaultdict
        from decimal import Decimal
        e={'OdNalogTip':'Stanica','OdNalogID':'S1','NaNalogTip':'Vozac','NaNalogID':'V1',
           'TipAmbalaze':'Gajba','Kolicina':'500'}
        acc=defaultdict(Decimal)
        _ledger_delta(e,acc)
        self.assertEqual(acc[('Stanica','S1','Gajba')],Decimal(-500))
        self.assertEqual(acc[('Vozac','V1','Gajba')],Decimal(500))
    def test_no_overwrite(self):
        with tempfile.TemporaryDirectory() as p:
            root=Path(p)
            with self.assertRaises(AuditError):cli(['--stage',str(root),'--expanded',str(root),
                '--seven',str(root),'--out',str(root)])


class Integration(unittest.TestCase):
    def test_full_7815_and_tamper(self):
        original=prior.cli
        called=[]
        def callback(args):
            result=original(args)
            if called:return result
            called.append(1)
            stage=Path(args[args.index('--stage')+1]);expanded=Path(args[args.index('--expanded')+1]);seven=Path(args[args.index('--out')+1]);out=seven.parent/'draft'
            argv=['--stage',str(stage),'--expanded',str(expanded),'--seven',str(seven),'--out',str(out)]
            self.assertEqual(cli(argv),0)
            s=json.loads((out/'agrix_amb10_dryrun_summary.json').read_text())
            self.assertEqual(s['format'],VERSION)
            self.assertEqual(s['physical_events_total'],4675)
            self.assertEqual(s['active_legacy_source_legs_covered'],7479)
            self.assertEqual(s['source_negative_balance_accounts_preserved'],json.loads((seven/'agrix_seven_hold_summary.json').read_text())['legacy_negative_account_types_unchanged'])
            self.assertEqual(s['unexplained_account_balance_residuals'],0)
            self.assertEqual(s['accounts_different_until_q00700_resolved'],2)
            self.assertEqual(s['historical_storno_source_rows_unreplayed'],335)
            self.assertFalse(s['import_approved'])
            self.assertTrue(s['no_canonical_amb_id_allocated'])
            records=[json.loads(x) for x in (out/'AMB10_DRYRUN_ALL_CASES_PRIVATE.jsonl').read_text().splitlines()]
            self.assertEqual(len(records),4675)
            self.assertEqual(s['target_shaped_candidate_count'],sum(bool(x['movement_preview']) for x in records))
            self.assertEqual(s['target_shaped_hold_count'],sum(bool(x['blocking_reasons']) for x in records))
            self.assertTrue(all(not x['import_approved'] for x in records))
            self.assertEqual(len((out/'AMB10_HISTORICAL_STORNO_HOLD_PRIVATE.jsonl').read_text().splitlines()),335)
            with self.assertRaises(AuditError):cli(argv)
            # Detect upstream evidence mutation on retry to a new destination.
            p=seven/'expanded_plus_held_physical_cases_PRIVATE.jsonl'
            original_data=p.read_text();p.write_text(original_data+'\n')
            argv[-1]=str(seven.parent/'draft_wrong')
            with self.assertRaises(AuditError):cli(argv)
            self.assertFalse(Path(argv[-1]).exists())
            p.write_text(original_data)
            return result
        with redirect_stdout(io.StringIO()),patch.object(previous_tests,'cli',new=callback):
            previous_tests.TestFullSynthetic('test_end_to_end_7815_and_refuse_overwrite_tamper').test_end_to_end_7815_and_refuse_overwrite_tamper()

if __name__=='__main__':unittest.main()
