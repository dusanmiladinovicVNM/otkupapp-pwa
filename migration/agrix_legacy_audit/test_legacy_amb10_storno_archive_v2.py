import io,json,tempfile,unittest
from pathlib import Path
from contextlib import redirect_stdout
from unittest.mock import patch
from legacy_audit import AuditError
import legacy_amb10_storno_archive_v2 as m
import legacy_amb10_storno_blueprint_v2 as old
import test_legacy_amb10_storno_blueprint_v2 as parent


def r(n,typ='Otkup',doc='OTK-1',role='Kooperant',eid='K1',direction='Izlaz',driver='',qty='20',storno='Da'):
    return parent.r(n,typ,doc,role,eid,direction,driver,qty,storno)

class Unit(unittest.TestCase):
 def test_no_driver_two_legs_storno_not_modified(self):
  pair={'source_row_ordinals':'1;2'}
  by_row={1:r(1,driver='V1'),2:r(2,role='Stanica',eid='S1',direction='Ulaz',driver='V1')}
  verdict,event=m._driver_pair_analysis(pair,by_row)
  self.assertEqual(verdict['class'],'NONCONSERVING_DERIVED_DRIVER_VECTOR')
  self.assertIsNone(event)
  self.assertEqual(by_row[1]['Stornirano'],'Da')
 def test_two_leg_driver_vector_nonconservation(self):
  pair={'source_row_ordinals':'1;2'}
  by_row={1:r(1,driver=''),2:r(2,role='Stanica',eid='S1',direction='Ulaz',driver='V1')}
  verdict,event=m._driver_pair_analysis(pair,by_row)
  self.assertEqual(verdict['class'],'NONCONSERVING_DERIVED_DRIVER_VECTOR')
  self.assertIsNone(event)
 def test_driver_vector_with_exact_two_account_reduction(self):
  # Two station legs with same driver and opposite directions: driver effects
  # cancel; the same station effects cancel as well -> not a nonzero transfer.
  pair={'source_row_ordinals':'1;2'}
  by_row={1:r(1,role='Stanica',eid='S1',direction='Izlaz',driver='V1'),
          2:r(2,role='Stanica',eid='S1',direction='Ulaz',driver='V1')}
  verdict,event=m._driver_pair_analysis(pair,by_row)
  self.assertEqual(verdict['class'],'MULTI_ACCOUNT_OR_NONUNIT_VECTOR')
  self.assertIsNone(event)
 def test_invalid_pair(self):
  pair={'source_row_ordinals':'1;1'}
  with self.assertRaises(AuditError):m._driver_pair_analysis(pair,{1:r(1)})
 def test_overwrite_prevented(self):
  with tempfile.TemporaryDirectory() as tmp:
   x=['--'+a for a in ('stage','expanded','seven','dryrun','blueprint','out')]
   args=[]
   for k in x:args.extend((k,tmp))
   with self.assertRaises(AuditError):m.cli(args)

class Integration(unittest.TestCase):
 def test_full_7815_storno_archive_and_tamper(self):
  previous=old.cli;called=[]
  def cb(args):
   status=previous(args)
   if called:return status
   called.append(True)
   folders={a:Path(args[args.index('--'+a)+1]) for a in ('stage','expanded','seven','dryrun','out')}
   dest=folders['out'].parent/'storno_archive'
   cmd=[]
   for n in ('stage','expanded','seven','dryrun'):
    cmd+=['--'+n,str(folders[n])]
   cmd+=['--blueprint',str(folders['out']),'--out',str(dest)]
   self.assertEqual(m.cli(cmd),0)
   rep=json.loads((dest/'agrix_storno_archive_summary.json').read_text())
   self.assertEqual(rep['historical_storno_rows_archived'],335)
   self.assertEqual(rep['previously_proven_source_rows'],335)
   self.assertEqual(rep['historical_rows_still_physically_unresolved'],0)
   self.assertEqual(rep['historical_originals_accounting_balance_effect'],0)
   self.assertFalse(rep['import_approved'])
   self.assertEqual(len((dest/'STORNO_ARCHIVE_ALL_CANCELLED_ORIGINALS_PRIVATE.jsonl').read_text().splitlines()),335)
   with self.assertRaises(AuditError):m.cli(cmd)
   p=folders['out']/'STORNO_HISTORY_HELD_PRIVATE.jsonl'
   original=p.read_text();p.write_text(original+'\n');cmd[-1]=str(dest.parent/'tamper_archive')
   with self.assertRaises(AuditError):m.cli(cmd)
   self.assertFalse(Path(cmd[-1]).exists())
   p.write_text(original)
   return status
  with redirect_stdout(io.StringIO()),patch.object(old,'cli',new=cb):
   parent.Integration('test_full_7815_and_upstream_tamper').test_full_7815_and_upstream_tamper()

if __name__=='__main__':unittest.main()
