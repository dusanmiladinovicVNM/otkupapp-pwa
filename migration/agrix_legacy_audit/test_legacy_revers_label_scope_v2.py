import hashlib
import json
import tempfile
import unittest
from pathlib import Path
from legacy_audit import AuditError
from legacy_revers_label_scope_v2 import physical_rescue, cli
from legacy_stage_v2 import SOURCE_SHA

MASTERS={
 'tblStanice':{'S1':(1,{'StanicaID':'S1'})},
 'tblKooperanti':{'K1':(1,{'KooperantID':'K1'})},
 'tblVozaci':{'V1':(1,{'VozacID':'V1'})}
}

def item(pos,typ='OM-Izlaz-Koop',d='REV-1',role='Kooperant',direction='Ulaz',owner='K1',qty='20',driver='',status='',flags='ORPHAN_DOCUMENT_ID;UNRESOLVED_DOCUMENT_ID',date='45500'):
 return {'AmbID':f'AMB-{pos:05d}','source_row_ordinal':pos,'DokumentTip':typ,'DokumentID':d,
         'Datum':date,'TipAmbalaze':'Gajba','Kolicina':qty,'EntitetTip':role,
         'EntitetID':owner,'Smer':direction,'VozacID':driver,'Stornirano':status,'forensic_flags':flags}

class PhysicalRecovery(unittest.TestCase):
 def pair(self):
  return [item(1),item(2,role='Stanica',direction='Izlaz',owner='S1')]
 def test_exact_two_leg(self):
  records,events=physical_rescue(self.pair(),MASTERS)
  self.assertEqual(len(events),1)
  self.assertEqual(records[0]['status'],'PHYSICAL_EVIDENCE_ONLY')
  self.assertEqual(events[0]['from_type'],'Stanica')
  self.assertEqual(events[0]['to_type'],'Kooperant')
 def test_revers_return(self):
  a=item(1,typ='OM-Ulaz-Koop',role='Kooperant',direction='Izlaz',flags='DOCUMENT_TYPE_REQUIRES_SEMANTIC_MAPPING')
  b=item(2,typ='OM-Ulaz-Koop',role='Stanica',direction='Ulaz',owner='S1',flags='DOCUMENT_TYPE_REQUIRES_SEMANTIC_MAPPING')
  _,events=physical_rescue([a,b],MASTERS)
  self.assertEqual((events[0]['from_type'],events[0]['to_type']),('Kooperant','Stanica'))
 def test_firma_single(self):
  r=item(1,typ='OM-Ulaz-Firma',role='Stanica',direction='Ulaz',owner='S1',driver='V1',flags='DOCUMENT_TYPE_REQUIRES_SEMANTIC_MAPPING')
  a,ev=physical_rescue([r],MASTERS)
  self.assertEqual(len(ev),1)
  self.assertEqual((ev[0]['from_type'],ev[0]['to_type']),('Vozac','Stanica'))
 def test_firma_driver_missing_hold(self):
  r=item(1,typ='OM-Izlaz-Firma',role='Stanica',direction='Izlaz',owner='S1',flags='DOCUMENT_TYPE_REQUIRES_SEMANTIC_MAPPING')
  a,ev=physical_rescue([r],MASTERS)
  self.assertEqual(len(ev),0)
  self.assertIn('DRIVER_NOT_PROVEN',a[0]['reasons'])
 def test_collision_scoped_date_station(self):
  a,b=self.pair();c=item(3,d='REV-1',qty='10');d=item(4,d='REV-1',role='Stanica',direction='Izlaz',owner='S1',qty='10')
  rec,ev=physical_rescue([a,b,c,d],MASTERS)
  self.assertFalse(ev)
  self.assertEqual(len(rec),2)
  self.assertTrue(all('MULTIPLE_TRANSFERS_SAME_LABEL_STATION_DAY' in x['reasons'] for x in rec))
 def test_diff_date_same_number_ok(self):
  a,b=self.pair();c=item(3,d='REV-1',date='45501');d=item(4,d='REV-1',role='Stanica',direction='Izlaz',owner='S1',date='45501')
  _,ev=physical_rescue([a,b,c,d],MASTERS)
  self.assertEqual(len(ev),2)
 def test_driver_on_koop_hold(self):
  a,b=self.pair(); a['VozacID']='V1'
  _,ev=physical_rescue([a,b],MASTERS)
  self.assertFalse(ev)
 def test_invalid_account_hold(self):
  a,b=self.pair(); a['EntitetID']='WRONG'
  _,ev=physical_rescue([a,b],MASTERS)
  self.assertFalse(ev)
 def test_raw_storno_hold(self):
  a,b=self.pair(); a['Stornirano']='Da';b['Stornirano']='Da'
  _,ev=physical_rescue([a,b],MASTERS)
  self.assertFalse(ev)

class FullSyntheticStage(unittest.TestCase):
 def test_real_shape_7815_regression(self):
  with tempfile.TemporaryDirectory() as tmp:
   root=Path(tmp);stage=root/'stage';qt=root/'quarantine';prov=root/'provenance'
   for path in [stage,qt,prov]:path.mkdir()
   masters={
     'SOURCE__tblOtkup':[],'SOURCE__tblOtpremnica':[],'tblPrijemnica':[],
     'tblKooperanti':[{'KooperantID':'K1'}],'tblStanice':[{'StanicaID':'S1'}],
     'tblKupci':[{'KupacID':'C1'}],'tblVozaci':[{'VozacID':'V1'}],
     'tblTipAmbalaze':[{'TipAmbalaze':'Gajba'}]
   }
   legs=[]
   def push(typ,doc,role,direction,owner,qty='20',driver='',status='',label=''):
    i=len(legs)+1
    legs.append(dict(AmbID=f'AMB-{i:05d}',Datum='45500',DokumentTIP=typ,
            DokumentID=doc,TipAmbalaze='Gajba',Kolicina=qty,Smer=direction,
            EntitetTip=role,EntitetID=owner,VozacID=driver,Stornirano=status))
   # 1702 OTK headers, 1702 reciprocal pair events. Last 128 are historical storno.
   for i in range(1702):
    doc=f'OTK-{i+1:04d}';sto='Da' if i>=1574 else ''
    masters['SOURCE__tblOtkup'].append(dict(OtkupID=doc,KooperantID='K1',StanicaID='S1',
         TipAmbalaze='Gajba',KolAmbalaze='20',KolAmbIzdata='20',Stornirano=sto))
    push('Otkup',doc,'Kooperant','Izlaz','K1',status=sto)
    push('Otkup',doc,'Stanica','Ulaz','S1',status=sto)
   # 1187 extra OM pairs are attached to those 1187 OTK headers (2x1187 rows).
   for i in range(1187):
    doc=f'OTK-{i+1:04d}'
    push('OM-Izlaz-Koop',doc,'Kooperant','Ulaz','K1')
    push('OM-Izlaz-Koop',doc,'Stanica','Izlaz','S1')
   for i in range(907):
    doc=f'OTP-{i+1:04d}'; psto='Da' if i==0 else ''
    masters['SOURCE__tblOtpremnica'].append(dict(OtpremnicaID=doc,StanicaID='S1',VozacID='V1',
             KolAmbalaze='20',TipAmbalaze='Gajba',Stornirano=psto))
    push('Otpremnica',doc,'Stanica','Izlaz','S1',driver='V1')
   for i in range(1033):
    doc=f'PRJ-{i+1:04d}';sto='Da' if i<79 else ''
    masters['tblPrijemnica'].append(dict(PrijemnicaID=doc,KupacID='C1',VozacID='V1',
                KolAmbalaze='20',TipAmbalaze='Gajba',Stornirano=sto))
    push('Prijemnica',doc,'Kupac','Ulaz','C1',driver='V1',status=sto)
   # 43 pair events with only business labels, no historical opaque ReversID.
   for i in range(18):
    label=f'RV-OUT-{i}'
    push('OM-Izlaz-Koop',label,'Kooperant','Ulaz','K1')
    push('OM-Izlaz-Koop',label,'Stanica','Izlaz','S1')
   for i in range(25):
    label=f'RV-IN-{i}'
    push('OM-Ulaz-Koop',label,'Kooperant','Izlaz','K1')
    push('OM-Ulaz-Koop',label,'Stanica','Ulaz','S1')
   for i in range(5):
    typ='OM-Izlaz-Firma' if i<2 else 'OM-Ulaz-Firma'
    push(typ,f'RV-F-{i}','Stanica','Izlaz' if i<2 else 'Ulaz','S1',driver='V1')
   for i in range(6):push('Kupci-Otpremnica',f'UNKNOWN-{i}','Kupac','Izlaz','C1',driver='V1')
   self.assertEqual(len(legs),7815)
   # Ensure Q00700 is one of the 907 OTP rows (case is never rescued).
   swap=3404+2374 # location of OTP-1, 0-based
   legs[699],legs[swap]=legs[swap],legs[699]
   # AmbID must remain unique and independent of the row ordinal.
   masters['UNCONVERTED__tblAmbalaza']=legs
   manifest={'source_sha256':SOURCE_SHA,'schema_sha256':'MIG05_2_SYNTH',
             'status':'STAGED_ONLY_NO_XLSM_CREATED',
             'canonical_stage_counts':{k:len(v) for k,v in masters.items()}}
   (stage/'agrix_stage_summary.json').write_text(json.dumps(manifest),encoding='utf-8')
   for tab,rows in masters.items():
    (stage/(tab+'.jsonl')).write_text(''.join(json.dumps(x)+'\n' for x in rows),encoding='utf-8')
   qt_summary={'format':'MIG-04.4','source_sha256':SOURCE_SHA,'schema_sha256':'MIG05_2_SYNTH','quarantined_active_rows':98}
   (qt/'agrix_amb10_quarantine_summary.json').write_text(json.dumps(qt_summary),encoding='utf-8')
   pr_summary={'format':'MIG-05.1','source_sha256':SOURCE_SHA,'schema_sha256':'MIG05_2_SYNTH',
          'om_quarantined_source_rows':91,'om_rows_with_revers_id':0,
          'all_om_provenance':{'BUSINESS_LABEL_ONLY_NO_OPAQUE_ID':91,'OTKUP_LINK_EXACT':2374}}
   (prov/'agrix_revers_provenance_summary.json').write_text(json.dumps(pr_summary),encoding='utf-8')
   srcfile=stage/'UNCONVERTED__tblAmbalaza.jsonl';before=hashlib.sha256(srcfile.read_bytes()).hexdigest()
   out=root/'output';self.assertEqual(cli(['--stage',str(stage),'--quarantine',str(qt),'--provenance',str(prov),'--out',str(out)]),0)
   result=json.loads((out/'agrix_revers_scope_summary.json').read_text(encoding='utf-8'))
   self.assertEqual(result['physical_candidate_events_by_shape'],{'PAIR_KOOP':43,'SINGLE_FIRMA':5})
   self.assertEqual(result['physical_candidate_source_rows'],91)
   self.assertEqual(result['revers_rows_still_on_hold'],0)
   self.assertEqual(result['expanded_shadow_event_count'],4669)
   self.assertEqual(result['still_quarantined_rows_total'],7)
   self.assertTrue(result['amb_00700_stays_on_hold'])
   self.assertEqual(before,hashlib.sha256(srcfile.read_bytes()).hexdigest())
   with self.assertRaises(AuditError):cli(['--stage',str(stage),'--quarantine',str(qt),'--provenance',str(prov),'--out',str(out)])

if __name__=='__main__':unittest.main()
