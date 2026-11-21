import json
import tempfile
import unittest
from collections import Counter
from pathlib import Path
from unittest.mock import patch

from legacy_audit import AuditError
import legacy_amb10_forensics_v2 as f
import legacy_amb10_shadow_v2 as shadow
import legacy_amb10_semantic_preview_v2 as sem
import legacy_revers_label_scope_v2 as old
from legacy_expanded_physical_stage_v2 import cli, _validate_report, VERSION
import test_legacy_revers_label_scope_v2 as oldtests


class Tests(unittest.TestCase):
    def test_mismatch_rejected(self):
        with self.assertRaises(AuditError):
            _validate_report({'format':'MIG-05.2','physical_candidate_events_total':48},
                             {'format':'MIG-05.2','physical_candidate_events_total':47},'old')

    def test_full_7815_synthetic_and_report_tamper(self):
        original=old.cli
        def wrap(args):
            # This fixture constructs 7815 rows with 91 historical Revers legs,
            # 335 STORNO, Q-00700, and 7 quarantined legs in tempfile.
            result=original(args)
            stage=Path(args[args.index('--stage')+1]);quarantine=Path(args[args.index('--quarantine')+1]);prov=Path(args[args.index('--provenance')+1]);scope=Path(args[args.index('--out')+1]);
            b,rows=f.analyze(stage);_,events,_,_=shadow.build_report(rows,b)
            _,_,idx=f.inputs(stage)
            proposal=[sem.classify_event(ev,rows,idx) for ev in events]
            semantic_dir=scope.parent/'semantic';semantic_dir.mkdir()
            semantic_summary={
                'format':'MIG-05.0','source_sha256':f.SOURCE_SHA,
                'schema_sha256':b['schema_sha256'],
                'physical_events_examined':len(events),
                'draft_evidence_status_counts':dict(Counter(x['evidence_status'] for x in proposal)),
                'proposed_business_movement_kind_counts':dict(Counter(x['VrstaKretanja_PROPOSAL'] for x in proposal)),
                'integrity_hold_reasons':{},
            }
            (semantic_dir/'agrix_amb10_semantic_preview_summary.json').write_text(json.dumps(semantic_summary))
            target=scope.parent/'expanded'
            params=['--stage',str(stage),'--quarantine',str(quarantine),
                    '--provenance',str(prov),'--revers-scope',str(scope),
                    '--semantic',str(semantic_dir),'--out',str(target)]
            self.assertEqual(cli(params),0)
            data=json.loads((target/'agrix_expanded_physical_summary.json').read_text())
            self.assertEqual(data['format'],VERSION)
            self.assertEqual(data['expanded_physical_events'],4669)
            self.assertEqual(data['physical_source_legs'],7473)
            self.assertEqual(data['still_excluded_active_source_legs'],7)
            self.assertEqual(data['historical_storno_rows_unconverted'],335)
            self.assertEqual(data['unexplained_account_model_residual_count'],0)
            self.assertTrue(data['q00700_remains_excluded'])
            self.assertEqual(data['remaining_hold_document_types'],{'Otpremnica':1,'Kupci-Otpremnica':6})
            self.assertEqual(sum(data['evidence_status_counts'].values()),4669)
            self.assertFalse(data['ready_for_import'])
            self.assertEqual(len((target/'expanded_physical_cases_PRIVATE.jsonl').read_text().splitlines()),4669)
            self.assertEqual(len((target/'seven_remaining_legacy_rows_PRIVATE.csv').read_text(encoding='utf-8-sig').splitlines()),8)
            # Refuse output overwrite; no files should be modified.
            with self.assertRaises(AuditError):cli(params)
            # A changed upstream summary must fail, even if counts otherwise look plausible.
            semantic_summary['proposed_business_movement_kind_counts']['UNASSIGNED']+=1
            (semantic_dir/'agrix_amb10_semantic_preview_summary.json').write_text(json.dumps(semantic_summary))
            params[-1]=str(scope.parent/'bad')
            with self.assertRaises(AuditError):cli(params)
            self.assertFalse((scope.parent/'bad').exists())
            return result
        # The old full end-to-end fixture checks its own invariants after our wrapper.
        # It calls CLI only once on fresh data and again to assert no overwrite.
        # Wrap only the FIRST invocation. The second must go directly to orig cli.
        calls=[]
        def once(args):
            calls.append(1)
            return wrap(args) if len(calls)==1 else original(args)
        with patch.object(oldtests,'cli',new=once):
            oldtests.FullSyntheticStage('test_real_shape_7815_regression').test_real_shape_7815_regression()

if __name__=='__main__':unittest.main()
