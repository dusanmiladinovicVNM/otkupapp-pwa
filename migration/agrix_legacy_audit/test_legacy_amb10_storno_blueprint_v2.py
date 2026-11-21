import io,json,tempfile,unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch
from legacy_audit import AuditError
import legacy_amb10_storno_blueprint_v2 as m
import test_legacy_amb10_canonical_dryrun_v2 as previous

MASTERS={
 'tblStanice':{'S1':1},'tblKooperanti':{'K1':1},'tblKupci':{'C1':1},
 'tblVozaci':{'V1':1},'tblTipAmbalaze':{'gajba':1}}

def r(n,typ='Otkup',doc='OTK-1',role='Kooperant',eid='K1',direction='Izlaz',driver='',qty='20',storno='Da'):
 return {'source_row_ordinal':n,'AmbID':f'AMB-{n:05d}','DokumentTip':typ,
   'DokumentID':doc,'Datum':'45500','TipAmbalaze':'Gajba','Kolicina':qty,
   'EntitetTip':role,'EntitetID':eid,'Smer':direction,'VozacID':driver,'Stornirano':storno}

class Unit(unittest.TestCase):
 def test_old_storno_pair_is_cancelled_original_not_inverse(self):
  a,b=r(1),r(2,role='Stanica',eid='S1',direction='Ulaz')
  ev,held,shapes,reasons=m.reconstruct([a,b],MASTERS)
  self.assertEqual((len(ev),len(held)),(1,0))
  self.assertEqual(ev[0]['source_row_ordinals'],'1;2')
  self.assertEqual(ev[0]['OdNalogTip'],'Kooperant')
  self.assertIsNone(ev[0]['new_StornoOd'])
  self.assertFalse(ev[0]['import_approved'])
  self.assertEqual(ev[0]['historical_cancellation_timestamp'],'UNPROVEN')
 def test_driver_inverse_cancelled_original(self):
  a=r(3,typ='Prijemnica',doc='PRJ-1',role='Kupac',eid='C1',direction='Ulaz',driver='V1')
  ev,held,*_=m.reconstruct([a],MASTERS)
  self.assertEqual((len(ev),len(held)),(1,0))
  self.assertEqual((ev[0]['OdNalogTip'],ev[0]['NaNalogTip']),('Vozac','Kupac'))
 def test_ambiguous_reused_business_label_is_held(self):
  a,b=r(1),r(2,role='Stanica',eid='S1',direction='Ulaz')
  c=r(3,doc='OTK-1',qty='30',role='Kooperant')
  ev,held,*_=m.reconstruct([a,b,c],MASTERS)
  self.assertFalse(ev)
  self.assertEqual(sum(len(x['source_row_ordinals'].split(';')) for x in held),3)
 def test_wrong_mirror_rejected(self):
  a,b=r(1),r(2,role='Stanica',eid='S1',direction='Izlaz')
  ev,held,*_=m.reconstruct([a,b],MASTERS)
  self.assertEqual(len(ev),0);self.assertEqual(len(held),1)
 def test_no_double_use(self):
  a=r(1);b=r(1,doc='OTHER')
  with self.assertRaises(AuditError):m.reconstruct([a,b],MASTERS)
 def test_active_not_accepted(self):
  with self.assertRaises(AuditError):m.reconstruct([r(1,storno='')],MASTERS)
 def test_no_overwrite(self):
  with tempfile.TemporaryDirectory() as t:
   with self.assertRaises(AuditError):m.cli(['--stage',t,'--expanded',t,'--seven',t,'--dryrun',t,'--out',t])

class Integration(unittest.TestCase):
 def test_full_7815_and_upstream_tamper(self):
  original=previous.cli
  called=[]
  def callback(args):
   result=original(args)
   if called:return result
   called.append(1)
   stage=Path(args[args.index('--stage')+1])
   expanded=Path(args[args.index('--expanded')+1])
   seven=Path(args[args.index('--seven')+1])
   dryrun=Path(args[args.index('--out')+1])
   dest=dryrun.parent/'storno_blueprint'
   cmd=['--stage',str(stage),'--expanded',str(expanded),'--seven',str(seven),'--dryrun',str(dryrun),'--out',str(dest)]
   self.assertEqual(m.cli(cmd),0)
   summ=json.loads((dest/'agrix_storno_blueprint_summary.json').read_text())
   self.assertEqual(summ['format'],m.VERSION)
   self.assertEqual(summ['historical_source_storno_legs'],335)
   self.assertEqual(summ['cancelled_physical_event_candidates'],207)
   self.assertEqual(summ['cancelled_physical_source_legs_covered'],335)
   self.assertEqual(summ['historical_source_legs_held'],0)
   self.assertEqual(summ['cancelled_physical_shape_counts'],{'DRIVER_SINGLE':79,'MIRROR_PAIR':128})
   self.assertEqual(summ['source_negative_balance_accounts_preserved'],summ['source_negative_balance_accounts_preserved'])
   self.assertFalse(summ['canonical_import_approved'])
   self.assertEqual(len((dest/'STORNO_PHYSICAL_CANCELLED_EVENTS_PRIVATE.jsonl').read_text().splitlines()),207)
   with self.assertRaises(AuditError):m.cli(cmd)
   file=dryrun/'AMB10_HISTORICAL_STORNO_HOLD_PRIVATE.jsonl';old=file.read_text();file.write_text(old+'\n')
   cmd[-1]=str(dest.parent/'tamper_storno')
   with self.assertRaises(AuditError):m.cli(cmd)
   self.assertFalse(Path(cmd[-1]).exists())
   file.write_text(old)
   return result
  with redirect_stdout(io.StringIO()),patch.object(previous,'cli',new=callback):
   previous.Integration('test_full_7815_and_tamper').test_full_7815_and_tamper()

if __name__=='__main__':unittest.main()
