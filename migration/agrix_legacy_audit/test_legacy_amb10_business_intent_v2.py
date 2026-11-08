import json,tempfile,unittest
from pathlib import Path
from legacy_audit import AuditError
import legacy_amb10_business_intent_v2 as m
import test_legacy_amb10_identity_bridge_v2 as fixture

class TestMIG061(unittest.TestCase):
 def base(self,kind,doc='OM-Izlaz-Koop',fr='Stanica',to='Kooperant',issued='HOLD_OWNERSHIP'):
  e=fixture.ev(kind)
  e.update(legacy_DokumentTIP=doc,OdNalogTip=fr,NaNalogTip=to,evidence_status=issued)
  i={'physical_case_id':e['physical_case_id'],'source_row_ordinals':e['source_row_ordinals'],
     'identity_status':m.bridge.CATEGORIES[1] if kind=='MIG05_ORIGINAL_PHYSICAL' else m.bridge.CATEGORIES[2],
     'verified_existing_DokumentTIP':'Otkup' if kind=='MIG05_ORIGINAL_PHYSICAL' else None,
     'verified_existing_DokumentID':'OTK-1' if kind=='MIG05_ORIGINAL_PHYSICAL' else None,
     'provisional_group_token':'MIG-GROUP-x' if kind!='MIG05_ORIGINAL_PHYSICAL' else None,
     'canonical_new_doc_id':None,'import_approved':False}
  d={'physical_case_id':e['physical_case_id'],'source_row_ordinals':e['source_row_ordinals'],
     'movement_preview':None,'blocking_reasons':['PURPOSE_OR_OWNERSHIP_NOT_PROVEN'], 'import_approved':False}
  return e,i,d
 def test_old_otkup_issuing_is_purpose_not_title(self):
  r=m.decision(*self.base('MIG05_ORIGINAL_PHYSICAL'))
  self.assertEqual(r['proposed_kind'],'IZDATA_PRAZNA')
  self.assertEqual(r['possible_foreign_return'],'VRACANJE_TUDJE_AMBALAZE_REQUIRES_INDEPENDENT_LIABILITY_EVIDENCE')
  self.assertFalse(r['import_approved']);self.assertIsNone(r['canonical_document_id'])
 def test_revers_four_directions(self):
  cases=[('OM-Izlaz-Koop','Stanica','Kooperant','IZDATA_PRAZNA'),
         ('OM-Ulaz-Koop','Kooperant','Stanica','POVRAT_PRAZNE'),
         ('OM-Izlaz-Firma','Stanica','Vozac','PRENOS_INTERNO'),
         ('OM-Ulaz-Firma','Vozac','Stanica','PRENOS_INTERNO')]
  for doc,fr,to,kind in cases:
   with self.subTest(doc=doc):
    e,i,d=self.base('MIG05_2_RECOVERED_REVERS_PHYSICAL',doc,fr,to)
    r=m.decision(e,i,d)
    self.assertEqual(r['proposed_kind'],kind)
    self.assertFalse(r['import_approved'])
 def test_revers_inconsistent_pair_blocked(self):
  e,i,d=self.base('MIG05_2_RECOVERED_REVERS_PHYSICAL','OM-Ulaz-Firma','Stanica','Vozac')
  with self.assertRaises(AuditError):m.decision(e,i,d)
 def test_existing_goods_proposal_preserved(self):
  e,i,d=self.base('MIG05_ORIGINAL_PHYSICAL')
  e['legacy_DokumentTIP']='Otkup';e['evidence_status']='DIRECTION_KIND_PROPOSED_ONLY';e['VrstaKretanja_PROPOSAL']='AMBALAZA_UZ_ROBU'
  i['identity_status']=m.bridge.CATEGORIES[0]
  d['blocking_reasons']=[];d['movement_preview']={'VrstaKretanja':'AMBALAZA_UZ_ROBU','DokumentTIP':'Otkup','DokumentID':'OTK-1'}
  r=m.decision(e,i,d)
  self.assertEqual(r['proposed_kind'],'AMBALAZA_UZ_ROBU')
 def test_unknown_stays_unassigned(self):
  e,i,d=self.base('MIG05_4_UNMODELLED_LEG_PHYSICAL_ONLY')
  i['identity_status']=m.bridge.CATEGORIES[3]
  r=m.decision(e,i,d)
  self.assertIsNone(r['proposed_kind'])
 def test_fail_closed_on_row_or_authorization(self):
  e,i,d=self.base('MIG05_ORIGINAL_PHYSICAL')
  i['source_row_ordinals']='x'
  with self.assertRaises(AuditError):m.decision(e,i,d)
  e,i,d=self.base('MIG05_ORIGINAL_PHYSICAL')
  d['import_approved']=True
  with self.assertRaises(AuditError):m.decision(e,i,d)
 def test_no_source_overwrite(self):
  with tempfile.TemporaryDirectory() as t:
   args=[]
   for k in ('stage','expanded','seven','dryrun','storno-archive','identity','out'):args+=['--'+k,t]
   with self.assertRaises(AuditError):m.cli(args)

class TestLargeSynthetic(unittest.TestCase):
 def test_4675_pinned_shapes_intent_proposals_no_autoapproval(self):
  obj=TestMIG061();all_rows=[]
  for n in range(4675):
   if n<3485:
    e,i,d=obj.base('MIG05_ORIGINAL_PHYSICAL')
    e['legacy_DokumentTIP']='Otkup';e['evidence_status']='DIRECTION_KIND_PROPOSED_ONLY';e['VrstaKretanja_PROPOSAL']='AMBALAZA_UZ_ROBU'
    i['identity_status']=m.bridge.CATEGORIES[0]
    d['blocking_reasons']=[];d['movement_preview']={'VrstaKretanja':'AMBALAZA_UZ_ROBU','DokumentTIP':'Otkup','DokumentID':'OTK-1'}
   elif n<4621:e,i,d=obj.base('MIG05_ORIGINAL_PHYSICAL')
   elif n<4669:
    vals=[('OM-Izlaz-Koop','Stanica','Kooperant'),('OM-Ulaz-Koop','Kooperant','Stanica'),('OM-Izlaz-Firma','Stanica','Vozac'),('OM-Ulaz-Firma','Vozac','Stanica')]
    doc,fr,to=vals[(n-4621)%4]
    e,i,d=obj.base('MIG05_2_RECOVERED_REVERS_PHYSICAL',doc,fr,to)
   else:
    e,i,d=obj.base('MIG05_4_UNMODELLED_LEG_PHYSICAL_ONLY');i['identity_status']=m.bridge.CATEGORIES[3]
   e['physical_case_id']=f'PHY-{n+1:05d}';i['physical_case_id']=e['physical_case_id'];d['physical_case_id']=e['physical_case_id']
   all_rows.append(m.decision(e,i,d))
  self.assertEqual(len(all_rows),4675)
  self.assertEqual(sum(bool(r['proposed_kind']) for r in all_rows),4669)
  self.assertEqual(sum(not r['proposed_kind'] for r in all_rows),6)
  self.assertTrue(all(r['approval_status']=='PROPOSAL_ONLY_NOT_CANONICAL' and not r['import_approved'] for r in all_rows))
  self.assertTrue(all(r['ownership_or_third_party_title']=='UNPROVEN' for r in all_rows))


class TestPipelineGuard(unittest.TestCase):
 def test_cli_mocked_4675_records_integrity_and_output(self):
  from unittest.mock import patch
  import hashlib
  with tempfile.TemporaryDirectory() as t:
   root=Path(t);dirs={k:root/k for k in ('stage','expanded','seven','dryrun','storno-archive','identity')}
   for p in dirs.values():p.mkdir()
   triples=[];fixture_obj=TestMIG061()
   for n in range(4675):
    if n<3485:
     e,i,d=fixture_obj.base('MIG05_ORIGINAL_PHYSICAL')
     e['legacy_DokumentTIP']='Otkup';e['evidence_status']='DIRECTION_KIND_PROPOSED_ONLY';e['VrstaKretanja_PROPOSAL']='AMBALAZA_UZ_ROBU'
     i['identity_status']=m.bridge.CATEGORIES[0];d['blocking_reasons']=[]
     d['movement_preview']={'VrstaKretanja':'AMBALAZA_UZ_ROBU','DokumentTIP':'Otkup','DokumentID':'OTK-1'}
    elif n<4621:e,i,d=fixture_obj.base('MIG05_ORIGINAL_PHYSICAL')
    elif n<4669:
     variants=[('OM-Izlaz-Koop','Stanica','Kooperant'),('OM-Ulaz-Koop','Kooperant','Stanica'),('OM-Izlaz-Firma','Stanica','Vozac'),('OM-Ulaz-Firma','Vozac','Stanica')]
     doc,fr,to=variants[(n-4621)%4];e,i,d=fixture_obj.base('MIG05_2_RECOVERED_REVERS_PHYSICAL',doc,fr,to)
    else:e,i,d=fixture_obj.base('MIG05_4_UNMODELLED_LEG_PHYSICAL_ONLY');i['identity_status']=m.bridge.CATEGORIES[3]
    for x in (e,i,d):x['physical_case_id']=f'PHY-{n+1:05d}'
    triples.append((e,i,d))
   events,identities,drys=zip(*triples)
   private=dirs['identity']/'AMB10_IDENTITY_BRIDGE_ALL_PRIVATE.jsonl'
   private.write_text(''.join(m.bridge.canon(v) for v in identities),encoding='utf-8',newline='\n')
   h=hashlib.sha256(private.read_bytes()).hexdigest()
   base={'source_sha256':m.f.SOURCE_SHA,'schema_sha256':'schema-synthetic',
         'legacy_source_negative_accounts_accepted':35,
         'remaining_events_without_existing_source_doc_pk':54}
   (dirs['identity']/'agrix_amb10_identity_bridge_summary.json').write_text(json.dumps(dict(base,identity_bridge_evidence_sha256=h)))
   # No stage used by the mocked upstream; this tests command composition,
   # identity file tamper gate, aggregate classifications, and no overwrite.
   argv=[]
   for k,v in dirs.items():argv+=['--'+k,str(v)]
   argv+=['--out',str(root/'result')]
   with patch.object(m.bridge,'build',return_value=(base,identities)),\
        patch.object(m.chrono,'_verify_inputs',return_value=(None,events,None)),\
        patch.object(m.bridge.dry,'build',return_value=(None,drys,[{'unexplained':'0'}],None)):
    self.assertEqual(m.cli(argv),0)
    result=json.loads((root/'result'/'agrix_business_intent_summary.json').read_text())
    self.assertEqual(result['active_physical_events'],4675)
    self.assertEqual(result['events_with_proposed_kind_not_approved'],4669)
    self.assertEqual(result['events_without_kind_proposal'],6)
    self.assertEqual(result['existing_otkup_issued_empty_intent_proposals'],1136)
    self.assertFalse(result['import_approved'])
    self.assertEqual(len((root/'result'/m.PRIVATE).read_text().splitlines()),4675)
    with self.assertRaises(AuditError):m.cli(argv)
    private.write_text(private.read_text()+'\n')
    argv[-1]=str(root/'tampered')
    with self.assertRaises(AuditError):m.cli(argv)
    self.assertFalse((root/'tampered').exists())

if __name__=="__main__":unittest.main()
