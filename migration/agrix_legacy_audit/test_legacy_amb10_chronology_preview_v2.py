import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from legacy_audit import AuditError
from legacy_amb10_chronology_preview_v2 import date_key,chronological_stress,cli,VERSION,_verify_inputs
import test_legacy_seven_hold_rescue_v2 as prior
import test_legacy_expanded_physical_stage_v2 as expanded_test

class DateTests(unittest.TestCase):
 def test_excel_serial(self):
  self.assertEqual(date_key('45500').isoformat(),'2024-07-27')
 def test_numeric_serial_with_fraction(self):
  self.assertEqual(date_key('45500.85').isoformat(),'2024-07-27')
 def test_local_date(self):
  self.assertEqual(date_key('27.07.2024').isoformat(),'2024-07-27')
 def test_iso(self):
  self.assertEqual(date_key('2024-07-27').isoformat(),'2024-07-27')
 def test_datetime_without_imaginary_precision(self):
  self.assertEqual(date_key('2024-07-27T11:15:20').isoformat(),'2024-07-27')
 def test_ambiguous_us_date(self):
  self.assertIsNone(date_key('07/08/2024'))
 def test_blank(self):
  self.assertIsNone(date_key(''))
 def test_invalid_calendar(self):
  self.assertIsNone(date_key('31.02.2024'))
 def test_invalid_serial(self):
  self.assertIsNone(date_key('NaN'))
 def test_unknown_date(self):
  self.assertIsNone(date_key('foo'))
 def test_no_input_dirs(self):
  with tempfile.TemporaryDirectory() as temp:
   with self.assertRaises(AuditError):_verify_inputs(Path(temp)/'a',Path(temp)/'b',Path(temp)/'c')

class NetActive(unittest.TestCase):
 def event(self,n,day,fr='Stanica',to='Vozac',q='20'):
  return {'physical_case_id':f'PHY-{n:05d}','source_row_ordinals':str(n),
    'legacy_Datum':day,'TipAmbalaze':'Gajba','Kolicina':q,
    'OdNalogTip':fr,'OdNalogID':'S1' if fr=='Stanica' else 'V1',
    'NaNalogTip':to,'NaNalogID':'V1' if to=='Vozac' else 'S1',
    'VrstaKretanja_PROPOSAL':'UNASSIGNED','evidence_status':'HOLD_OWNERSHIP',
    'event_provenance':'MIG05_ORIGINAL_PHYSICAL','ownership':'UNPROVEN',
    'opening_balance':'UNPROVEN','historical_storno':'NOT_REPLAYED',
    'canonical_AMB10_import_authorized':'NO'}
 def test_invariant_rejects_source_repeat(self):
  events=[self.event(i,'45500') for i in range(1,4676)]
  events[1]['source_row_ordinals']='1'
  with self.assertRaisesRegex(AuditError,'lineage partition'):
   chronological_stress(events,{i:'45500' for i in range(1,4676)})
 def test_authorization_rejected(self):
  e=self.event(1,'45500');e['canonical_AMB10_import_authorized']='YES'
  with self.assertRaisesRegex(AuditError,'approval'):
   chronological_stress([e],{1:'45500'})
 def test_unknown_date_count(self):
  events=[self.event(i,'45500') for i in range(1,4676)]
  for i in range(2804):events[i]['source_row_ordinals'] += ';'+str(4676+i)
  events[14]['legacy_Datum']='?'
  summary,details,accounts=chronological_stress(events,{i:'45500' for i in range(1,7480)})
  self.assertEqual(summary['date_unknown_event_count'],1)
  self.assertTrue(summary['account_type_balances_with_multiple_events_same_day'])
  self.assertEqual(summary['physical_legacy_source_legs'],7479)
  self.assertEqual(summary['hypothetical_best_day_order_accounts_with_zero_opening_deficit'],1)
  self.assertTrue(summary['hypothetical_worst_day_order_accounts_with_zero_opening_deficit']>=1)
  self.assertEqual(details[14]['date_status'],'DATE_CONFLICT_OR_UNKNOWN')
  self.assertEqual(sum(a['date_unknown_events'] for a in accounts),2)

class Integration(unittest.TestCase):
 def test_existing_synthetic_7815(self):
  original=prior.cli
  fired=[]
  def hook(args):
   result=original(args)
   if fired:return result
   fired.append(True)
   stage=Path(args[args.index('--stage')+1]);expanded=Path(args[args.index('--expanded')+1]);seven=Path(args[args.index('--out')+1]);out=seven.parent/'chrono'
   self.assertEqual(cli(['--stage',str(stage),'--expanded',str(expanded),'--seven',str(seven),'--out',str(out)]),0)
   s=json.loads((out/'agrix_amb10_chronology_summary.json').read_text())
   self.assertEqual(s['format'],VERSION)
   self.assertEqual(s['physical_case_count'],4675)
   self.assertEqual(s['physical_legacy_source_legs'],7479)
   self.assertEqual(s['historical_storno_rows_not_replayed'],335)
   self.assertFalse(s['canonical_amb10_import_approved'])
   self.assertEqual(len((out/'chronology_cases_PRIVATE.csv').read_text(encoding='utf-8-sig').splitlines()),4676)
   with self.assertRaises(AuditError):cli(['--stage',str(stage),'--expanded',str(expanded),'--seven',str(seven),'--out',str(out)])
   p=seven/'expanded_plus_held_physical_cases_PRIVATE.jsonl'
   prior_text=p.read_text();p.write_text(prior_text.replace('"historical_storno": "NOT_REPLAYED"','"historical_storno": "REPLAYED"',1))
   with self.assertRaises(AuditError):_verify_inputs(stage,expanded,seven)
   p.write_text(prior_text)
   return result
  with patch.object(prior,'cli',new=hook):
   prior.TestFullSynthetic('test_end_to_end_7815_and_refuse_overwrite_tamper').test_end_to_end_7815_and_refuse_overwrite_tamper()

if __name__=='__main__':unittest.main()
