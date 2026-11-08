import json
import hashlib
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import patch

from legacy_audit import AuditError
import legacy_amb10_storno_lineage_v2 as m
import legacy_amb10_forensics_v2 as f


def ambrow(n, doc='T1', doctype='Otpremnica', direction='Izlaz', qty='10', storno='', driver='D1'):
    return {'AmbID':f'AMB-{n:05d}','Datum':'45200','TipAmbalaze':'T',
            'Kolicina':qty,'Smer':direction,'EntitetTip':'Stanica','EntitetID':'S1',
            'VozacID':driver,'DokumentID':doc,'DokumentTIP':doctype,
            'Stornirano':storno}


def fixture_rows(parentstorno='Da', legstorno='', inverse_doc=None):
    rows=[ambrow(i,doc='O1',doctype='Otkup',driver='') for i in range(1,7816)]
    rows[699]=ambrow(700,qty='500',storno=legstorno)
    if inverse_doc:
        rows[700]=ambrow(701,doc=inverse_doc,direction='Ulaz',qty='500')
    return rows


def synth_stage(folder, rows, parentstorno='Da'):
    stages={
      'UNCONVERTED__tblAmbalaza':rows,
      'SOURCE__tblOtkup':[{'OtkupID':'O1','KooperantID':'K1','StanicaID':'S1','KolAmbalaze':'10',
                            'KolAmbIzdata':'0','Stornirano':''}],
      'SOURCE__tblOtpremnica':[
        {'OtpremnicaID':'T1','StanicaID':'S1','VozacID':'D1','KolAmbalaze':'500','Stornirano':parentstorno},
        {'OtpremnicaID':'T2','StanicaID':'S1','VozacID':'D1','KolAmbalaze':'500','Stornirano':''},
      ],
      'tblPrijemnica':[{'PrijemnicaID':'P1','KupacID':'C1','VozacID':'D1','Stornirano':''}],
      'tblKooperanti':[{'KooperantID':'K1'}],
      'tblStanice':[{'StanicaID':'S1'}],
      'tblKupci':[{'KupacID':'C1'}],
      'tblVozaci':[{'VozacID':'D1'}],
      'tblTipAmbalaze':[{'TipAmbalaze':'T'}],
      'UNCONVERTED__tblStornoVeze':[{'StornoID':'V1','OldDocID':'T1'}],
    }
    mf={'format':'MIG-02.0','source_sha256':f.SOURCE_SHA,
        'status':'STAGED_ONLY_NO_XLSM_CREATED','schema_sha256':'SYNTHETIC',
        'canonical_stage_counts':{k:len(v) for k,v in stages.items()}}
    (folder/'agrix_stage_summary.json').write_text(json.dumps(mf),encoding='utf-8')
    for table,items in stages.items():
        (folder/(table+'.jsonl')).write_text(''.join(json.dumps(z)+'\n' for z in items),encoding='utf-8')
    return mf


class StornoLineageTests(TestCase):
    def test_storno_parent_active_leg_is_conflict_qty_match(self):
        with TemporaryDirectory() as tmp:
            p=Path(tmp);synth_stage(p,fixture_rows())
            rep,review,candidates,focus_row,focus,refs,_,_=m.build(p)
            self.assertEqual(rep['focus']['amb_state'],'ACTIVE')
            self.assertEqual(rep['focus']['parent_state'],'STORNO')
            self.assertTrue(rep['focus']['parent_amb_quantity_exact_match'])
            self.assertTrue(rep['focus']['parent_station_matches'])
            self.assertTrue(rep['focus']['parent_driver_matches'])
            self.assertTrue(rep['focus']['id_and_ordinal_same_source_record'])
            self.assertEqual(focus['issue_class'],'PARENT_STORNO_AMB_ACTIVE')
            self.assertEqual(rep['linked_storno_document_counts']['DOCUMENT_LEG_STATE_MISMATCH'],1)
            self.assertEqual(rep['focus']['literal_matches_in_stornoveze'],1)
            self.assertFalse(rep['focus']['same_document_exact_active_physical_offset_candidates'])
            self.assertEqual(candidates,[])

    def test_same_doc_opposite_active_is_candidate_not_proof(self):
        with TemporaryDirectory() as tmp:
            p=Path(tmp);synth_stage(p,fixture_rows(inverse_doc='T1'))
            rep,review,candidates,*_=m.build(p)
            self.assertEqual(rep['focus']['same_document_exact_active_physical_offset_candidates'],1)
            self.assertTrue(any(c['same_document_id_and_type']=='YES' for c in candidates))
            self.assertEqual(rep['focus']['approval'],'NONE')

    def test_other_doc_inverse_cannot_be_called_storno(self):
        with TemporaryDirectory() as tmp:
            p=Path(tmp);synth_stage(p,fixture_rows(inverse_doc='T2'))
            rep,review,candidates,*_=m.build(p)
            self.assertEqual(rep['focus']['same_document_exact_active_physical_offset_candidates'],0)
            self.assertEqual(rep['focus']['other_document_active_physical_offset_candidates'],1)
            self.assertEqual(candidates[0]['relation'],'PHYSICAL_OFFSET_CANDIDATE_NOT_STORNO_PROOF')

    def test_both_storno_historical_not_state_conflict(self):
        with TemporaryDirectory() as tmp:
            p=Path(tmp);synth_stage(p,fixture_rows(legstorno='Da'))
            rep,review,candidates,*_=m.build(p)
            self.assertEqual(rep['focus']['amb_state'],'STORNO')
            self.assertEqual(rep['linked_storno_document_counts']['DOCUMENT_LEG_STATE_MISMATCH'],0)
            self.assertEqual(rep['linked_storno_document_counts']['ACTIVE_AMB_STORNO_PARENT'],0)
            self.assertEqual(rep['review_cases'],1)

    def test_active_parent_storno_leg_conflict(self):
        with TemporaryDirectory() as tmp:
            p=Path(tmp);synth_stage(p,fixture_rows(legstorno='Da'),parentstorno='')
            rep,review,candidates,*_=m.build(p)
            self.assertEqual(rep['linked_storno_document_counts']['STORNO_AMB_ACTIVE_PARENT'],1)
            self.assertEqual(rep['focus']['amb_state'],'STORNO')

    def test_wrong_station_driver_context_is_visible_not_approved(self):
        with TemporaryDirectory() as tmp:
            p=Path(tmp);rows=fixture_rows();rows[699]['EntitetID']='S_WRONG';rows[699]['VozacID']='D_WRONG';synth_stage(p,rows)
            rep,review,candidates,*_=m.build(p)
            self.assertFalse(rep['focus']['parent_station_matches'])
            self.assertFalse(rep['focus']['parent_driver_matches'])
            self.assertEqual(rep['focus']['approval'],'NONE')

    def test_unknown_storno_token_fails_closed(self):
        with TemporaryDirectory() as tmp:
            p=Path(tmp);rows=fixture_rows();rows[699]['Stornirano']='NEVERHECKED';synth_stage(p,rows)
            with self.assertRaisesRegex(AuditError,'Unexpected Stornirano token'):
                m.build(p)

    def test_end_to_end_cli_no_edit_and_existing_output_guard(self):
        with TemporaryDirectory() as tmp:
            p=Path(tmp);stage=p/'stage';stage.mkdir();out=p/'output'
            synth_stage(stage,fixture_rows(inverse_doc='T1'))
            stage_hash=hashlib.sha256((stage/'UNCONVERTED__tblAmbalaza.jsonl').read_bytes()).hexdigest()
            self.assertEqual(m.cli(['--stage',str(stage),'--out',str(out)]),0)
            summary=json.loads((out/'agrix_storno_lineage_summary.json').read_text(encoding='utf-8'))
            self.assertEqual(summary['source_amb_rows'],7815)
            self.assertTrue((out/'Q00700_local_evidence_PRIVATE.txt').is_file())
            self.assertEqual(hashlib.sha256((stage/'UNCONVERTED__tblAmbalaza.jsonl').read_bytes()).hexdigest(),stage_hash)
            with self.assertRaisesRegex(AuditError,'already exists'):
                m.cli(['--stage',str(stage),'--out',str(out)])

    def test_id_row_mismatch_is_reported_and_not_equated(self):
        with TemporaryDirectory() as tmp:
            p=Path(tmp);rows=fixture_rows();rows[699]['AmbID']='AMB-00999';rows[998]['AmbID']='AMB-00700';synth_stage(p,rows)
            rep,*_=m.build(p)
            self.assertEqual(rep['focus']['focus_amb_id_found_at_ordinal'],999)
            self.assertFalse(rep['focus']['id_and_ordinal_same_source_record'])
