import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from legacy_audit import AuditError
from legacy_seven_hold_rescue_v2 import (physical_vector_only,load_prior_evidence,
    checked_summary,cli,VERSION)
import legacy_expanded_physical_stage_v2 as xp
import test_legacy_expanded_physical_stage_v2 as extests

MASTERS={
 'tblStanice':{'S1':True},'tblKupci':{'C1':True},
 'tblVozaci':{'V1':True},'tblTipAmbalaze':{'gajba':True}
}

def item(t='Kupac',s='Izlaz',who='C1',driver='V1',qty='25'):
 return {'source_row_ordinal':4,'AmbID':'AMB-4','DokumentTip':'weird-old-type',
    'DokumentID':'','Datum':'45500','TipAmbalaze':'Gajba','Kolicina':qty,
    'Smer':s,'EntitetTip':t,'EntitetID':who,'VozacID':driver,'Stornirano':''}

class TestVector(unittest.TestCase):
 def test_kupac_driver_direction(self):
  a,why=physical_vector_only(item(),MASTERS)
  self.assertFalse(why)
  self.assertEqual((a['from_type'],a['to_type']),('Kupac','Vozac'))
  self.assertEqual(a['document_id'],'')
 def test_kupac_ulaz(self):
  a,why=physical_vector_only(item(s='Ulaz'),MASTERS)
  self.assertFalse(why)
  self.assertEqual((a['from_type'],a['to_type']),('Vozac','Kupac'))
 def test_stanica_driver(self):
  a,why=physical_vector_only(item(t='Stanica',who='S1'),MASTERS)
  self.assertFalse(why)
  self.assertEqual((a['from_type'],a['to_type']),('Stanica','Vozac'))
 def test_koop_single_no_derived_driver(self):
  a,why=physical_vector_only(item(t='Kooperant'),MASTERS)
  self.assertIsNone(a);self.assertIn('NOT_DRIVER_INVERSE_ACCOUNT_CLASS',why)
 def test_bad_driver(self):
  a,why=physical_vector_only(item(driver='NOT_REGISTERED'),MASTERS)
  self.assertIsNone(a);self.assertIn('UNKNOWN_OR_MISSING_DRIVER',why)
 def test_bad_amount(self):
  a,why=physical_vector_only(item(qty='0'),MASTERS)
  self.assertIsNone(a);self.assertIn('INVALID_QUANTITY',why)
 def test_no_date(self):
  r=item();r['Datum']=''
  a,why=physical_vector_only(r,MASTERS)
  self.assertIsNone(a);self.assertIn('MISSING_DATE',why)
 def test_bad_account(self):
  a,why=physical_vector_only(item(who='C9'),MASTERS)
  self.assertIsNone(a);self.assertIn('UNKNOWN_SOURCE_ACCOUNT',why)
 def test_bad_pack_type(self):
  r=item();r['TipAmbalaze']='Fake'
  a,why=physical_vector_only(r,MASTERS)
  self.assertIsNone(a);self.assertIn('UNKNOWN_PACKAGING_TYPE',why)
 def test_storno_row_no_replay(self):
  r=item();r['Stornirano']='Da'
  a,why=physical_vector_only(r,MASTERS)
  self.assertIsNone(a);self.assertIn('NOT_ACTIVE',why)
 def test_no_prior_report(self):
  with tempfile.TemporaryDirectory() as d:
   with self.assertRaises(AuditError):checked_summary(Path(d),{'schema_sha256':'X'})

class TestFullSynthetic(unittest.TestCase):
 def test_end_to_end_7815_and_refuse_overwrite_tamper(self):
  original=extests.cli
  self_test=self
  invocations=[]
  def wrapped(args):
   invocations.append(1)
   result=original(args)
   if len(invocations)!=1:return result
   stage=Path(args[args.index('--stage')+1]);path=Path(args[args.index('--out')+1])
   target=path.parent/'seven'
   params=['--stage',str(stage),'--expanded',str(path),'--out',str(target)]
   self_test.assertEqual(cli(params),0)
   data=json.loads((target/'agrix_seven_hold_summary.json').read_text())
   self_test.assertEqual(data['format'],VERSION)
   self_test.assertEqual(data['previous_physical_events'],4669)
   self_test.assertEqual(data['recovered_extra_physical_vector_cases'],6)
   self_test.assertEqual(data['new_physical_events_total'],4675)
   self_test.assertEqual(data['new_physical_source_legs'],7479)
   self_test.assertEqual(data['remaining_active_hold_source_ordinals'],[700])
   self_test.assertEqual(data['unexplained_account_residual_count'],0)
   self_test.assertEqual(data['historical_storno_rows_unconverted'],335)
   self_test.assertFalse(data['any_import_approved'])
   self_test.assertEqual(len((target/'expanded_plus_held_physical_cases_PRIVATE.jsonl').read_text().splitlines()),4675)
   with self_test.assertRaises(AuditError):cli(params)
   # Public metadata cannot be edited after the projection was generated.
   prior=path/'agrix_expanded_physical_summary.json'
   old=prior.read_text();mod=json.loads(old);mod['physical_source_legs']=7472
   prior.write_text(json.dumps(mod))
   params[-1]=str(path.parent/'tampered')
   with self_test.assertRaises(AuditError):cli(params)
   self_test.assertFalse((path.parent/'tampered').exists())
   prior.write_text(old)
   # Private event evidence checksum must match MIG-05.3.
   evidence=path/'expanded_physical_cases_PRIVATE.jsonl'
   content=evidence.read_text();evidence.write_text(content+'\n')
   with self_test.assertRaises(Exception):cli(params)
   evidence.write_text(content)
   return result
  # Reuse full 7,815-row synthetic fixture from MIG-05.3 regression.
  with patch.object(extests,'cli',new=wrapped):
   extests.Tests('test_full_7815_synthetic_and_report_tamper').test_full_7815_synthetic_and_report_tamper()

if __name__=='__main__':unittest.main()
