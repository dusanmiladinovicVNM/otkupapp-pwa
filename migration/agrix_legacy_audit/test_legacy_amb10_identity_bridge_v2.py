import hashlib,json,tempfile,unittest
from pathlib import Path
from legacy_audit import AuditError
import legacy_amb10_identity_bridge_v2 as m


def ev(kind='MIG05_ORIGINAL_PHYSICAL'):
 return {'physical_case_id':'PHY-00001','source_row_ordinals':'1;2','source_amb_ids':'AMB1;AMB2',
         'legacy_DokumentTIP':'OM-Izlaz-Koop','legacy_DokumentID':'OTK-1','legacy_Datum':'45500',
         'TipAmbalaze':'GAJBA','Kolicina':'20','OdNalogTip':'Stanica','OdNalogID':'S1',
         'NaNalogTip':'Kooperant','NaNalogID':'K1','event_provenance':kind,
         'evidence_status':'HOLD_OWNERSHIP','VrstaKretanja_PROPOSAL':'UNASSIGNED',
         'canonical_AMB10_import_authorized':'NO'}

def legs():
 return [{'AmbID':'AMB1','DokumentTip':'OM-Izlaz-Koop','DokumentID':'OTK-1',
          'EntitetTip':'Stanica','EntitetID':'S1','Smer':'Izlaz','Stornirano':'',
          'Kolicina':'20','TipAmbalaze':'GAJBA','Datum':'45500'},
         {'AmbID':'AMB2','DokumentTip':'OM-Izlaz-Koop','DokumentID':'OTK-1',
          'EntitetTip':'Kooperant','EntitetID':'K1','Smer':'Ulaz','Stornirano':'',
          'Kolicina':'20','TipAmbalaze':'GAJBA','Datum':'45500'}]

class Tests(unittest.TestCase):
 def setUp(self):
  self.index={'SOURCE__tblOtkup':{'OTK-1':(1,{'Stornirano':'','StanicaID':'S1','KooperantID':'K1','KolAmbIzdata':'20','TipAmbalaze':'GAJBA'})}}
 def test_otkup_existing_identity_proven_not_kind(self):
  pair=ev();a=m.identity_link(pair,dict(physical_case_id='PHY-00001',source_row_ordinals='1;2',
     movement_preview=None,blocking_reasons=['CANONICAL_SOURCE_DOCUMENT_IDENTITY_UNRESOLVED'],
     import_approved=False),legs(),self.index,set())
  self.assertEqual(a['verified_existing_DokumentID'],'OTK-1')
  self.assertEqual(a['identity_status'],m.CATEGORIES[1])
  self.assertEqual(a['purpose_status'],'ISSUED_OR_FOREIGN_RETURN_UNPROVEN')
  self.assertIsNone(a['canonical_new_doc_id'])
  self.assertFalse(a['import_approved'])
 def test_otkup_cannot_pass_if_quantity_or_identity_disagree(self):
  x=ev();row=legs();row[0]['Kolicina']='19'
  with self.assertRaises(AuditError):m.prove_existing_otkup(x,row,self.index)
  row=legs();row[0]['EntitetID']='wrong'
  with self.assertRaises(AuditError):m.prove_existing_otkup(x,row,self.index)
  x['legacy_DokumentID']='other'
  with self.assertRaises(AuditError):m.prove_existing_otkup(x,legs(),self.index)
 def test_goods_existing_parent(self):
  x=ev();x['legacy_DokumentTIP']='Otkup';x['event_provenance']='MIG05_ORIGINAL_PHYSICAL';x['evidence_status']='DIRECTION_KIND_PROPOSED_ONLY';x['VrstaKretanja_PROPOSAL']='AMBALAZA_UZ_ROBU'
  old={'physical_case_id':'PHY-00001','source_row_ordinals':'1;2','import_approved':False,'blocking_reasons':[],
       'movement_preview':{'DokumentTIP':'Otkup','DokumentID':'OTK-1','VrstaKretanja':'AMBALAZA_UZ_ROBU'}}
  a=m.identity_link(x,old,legs(),self.index,set())
  self.assertEqual(a['identity_status'],m.CATEGORIES[0])
  self.assertEqual(a['purpose_status'],'KIND_PROPOSED_NOT_APPROVED')
 def test_revers_scope_is_not_opaque_document_identity(self):
  x=ev('MIG05_2_RECOVERED_REVERS_PHYSICAL')
  x['legacy_DokumentID']='REV-12'
  rows=legs()
  for r in rows:r['DokumentID']='REV-12'
  scope=set();a=m.identity_link(x,{'physical_case_id':'PHY-00001','source_row_ordinals':'1;2',
      'import_approved':False,'movement_preview':None},rows,self.index,scope)
  self.assertEqual(a['identity_status'],m.CATEGORIES[2])
  self.assertIsNone(a['canonical_new_doc_id'])
  self.assertTrue(a['provisional_group_token'].startswith('MIG-GROUP-'))
  with self.assertRaises(AuditError):m.scope_revers(x,rows,scope)
 def test_unknown_document_not_silently_mapped(self):
  x=ev('MIG05_4_UNMODELLED_LEG_PHYSICAL_ONLY');x['source_row_ordinals']='1';x['source_amb_ids']='AMB1'
  x['legacy_DokumentTIP']='unregistered';x['legacy_DokumentID']=''
  rec=m.identity_link(x,{'physical_case_id':'PHY-00001','source_row_ordinals':'1',
     'movement_preview':None,'import_approved':False},legs(),self.index,set())
  self.assertEqual(rec['identity_status'],m.CATEGORIES[3])
  self.assertIsNone(rec['verified_existing_DokumentTIP'])
 def test_private_exact_evidence_guard(self):
  with tempfile.TemporaryDirectory() as t:
   p=Path(t)/'source.jsonl'
   p.write_text(m.canon({'x':1})+m.canon({'x':2}),encoding='utf-8')
   self.assertEqual(m.assert_records(p,[{'x':1},{'x':2}]),hashlib.sha256(p.read_bytes()).hexdigest())
   p.write_text(m.canon({'x':1}),encoding='utf-8')
   with self.assertRaises(AuditError):m.assert_records(p,[{'x':1},{'x':2}])
 def test_outdir_clobber_refused(self):
  with tempfile.TemporaryDirectory() as t:
   args=[]
   for x in ('stage','expanded','seven','dryrun','storno-archive','out'):args+=['--'+x,t]
   with self.assertRaises(AuditError):m.cli(args)

if __name__=='__main__': unittest.main()
