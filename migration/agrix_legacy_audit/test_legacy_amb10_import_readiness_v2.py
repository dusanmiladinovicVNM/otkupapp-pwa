import json,tempfile,unittest,hashlib
from pathlib import Path
from unittest.mock import patch
from legacy_audit import AuditError
import legacy_amb10_import_readiness_v2 as m

class TestMIG064(unittest.TestCase):
 def fixture(self,root):
  root.mkdir(parents=True,exist_ok=True)
  rows=[];headers=[];holds=[];num=iter(n for n in range(1,7481) if n!=700)
  for i in range(4675):
   kind=('AMBALAZA_UZ_ROBU' if i<3349 else
      'POVRAT_PRAZNE' if i<3485 else
      'IZDATA_PRAZNA' if i<4621 else
      'IZDATA_PRAZNA' if i<4639 else
      'POVRAT_PRAZNE' if i<4664 else
      'PRENOS_INTERNO' if i<4669 else None)
   cat=('EXISTING_GOODS_PARENT_AND_KIND_PROPOSED' if i<3485 else
        'EXISTING_OTKUP_PARENT_KIND_OWNERSHIP_HELD' if i<4621 else
        'SCOPED_LEGACY_REVERS_NEW_DOCUMENT_NEEDED' if i<4669 else None)
   ords=str(next(num))+((';'+str(next(num))) if i<2804 else '')
   cid='PHY-'+str(i)
   if cat is None:
    holds.append({'physical_case_id':cid,'source_row_ordinals':ords,
         'reason':'NO_SOURCE_DOCUMENT_ID_AND_NO_PROVEN_KIND','import_approved':False})
    continue
   if kind=='IZDATA_PRAZNA':fr,to='Stanica','Kooperant'
   elif kind=='POVRAT_PRAZNE':fr,to='Kooperant','Stanica'
   elif kind=='PRENOS_INTERNO':fr,to='Stanica','Vozac'
   else:fr,to='Kooperant','Stanica'
   rev=cat=='SCOPED_LEGACY_REVERS_NEW_DOCUMENT_NEEDED';docid='ADK-'+str(i) if rev else 'OTK-'+str(i)
   if rev:
    headers.append({'target_header_draft':{'AmbDokID':docid,'Vrsta':'REVERS',
      'BrojOwnerTip':'Stanica','BrojOwnerID':'ST','BrojDokumenta':'N'+str(i),'Datum':'2026-08-01'},
      'ownership_proven':False,'import_approved':False})
   blockers=['MIGRATION_ONLY_NOT_IMPORTABLE','LEGAL_TITLE_NOT_PROVEN',
     'HISTORICAL_CANCELLATIONS_NOT_REPLAYED','Q00700_PENDING']
   if kind=='IZDATA_PRAZNA':blockers.append('FOREIGN_PACKAGING_LIABILITY_NOT_PROVEN')
   if rev:blockers.append('NEW_REVERS_HEADER_NOT_APPROVED')
   p={k:None for k in m.linked.CANONICAL_COLUMNS}
   p.update(Datum='2026-08-01',TipAmbalaze='GAJBA',Kolicina='3',
        DokumentTIP='AmbalazaDokument' if rev else 'Otkup',DokumentID=docid,
        OdNalogTip=fr,OdNalogID='ST' if fr=='Stanica' else 'KO' if fr=='Kooperant' else 'VZ',
        NaNalogTip=to,NaNalogID='ST' if to=='Stanica' else 'KO' if to=='Kooperant' else 'VZ',
        VrstaKretanja=kind)
   rows.append({'physical_case_id':cid,'source_row_ordinals':ords,
    'identity_status':cat,'import_blockers':blockers,
    'proposed_target_columns':p,'import_approved':False,'canonical_AmbID_allocated':False})
  self.assertEqual((len(rows),len(holds),len(headers)),(4669,6,48))
  self.assertEqual(len({int(v) for rec in rows+holds for v in rec['source_row_ordinals'].split(';')}),7479)
  obj={'format':'MIG-06.3','status':'PRIVATE_LINKED_TARGET_SHAPE_PREVIEW_NO_IMPORT',
      'source_sha256':m.SOURCE_SHA,'schema_sha256':m.SCHEMA_SHA,
      'total_physical_events':4675,'target_shaped_proposals_not_approved':4669,
      'unknown_document_event_holds':6,'linked_new_revers_header_drafts':48,
      'new_revers_rows_linked':48,'rows_with_existing_parent':4621,
      'inv04_colliding_keys':0,'inv10_documents_multiple_party_pairs':0,
      'proposal_rows_with_inv04_or_inv10_hold':0,'source_active_legs_covered':7479,
      'historical_cancelled_legacy_rows_archived':335,
      'historical_cancelled_legacy_rows_automatically_replayed':0,
      'source_negative_account_types_allowed':35,'unexplained_account_balance_residuals':0,
      'approved_canonical_amb_ids':0,'synthetic_opening_balances':0,
      'foreign_liability_or_deficit_cover_generated':0,
      'active_q00700_held':True,'import_approved':False,'target_xlsm_created_or_modified':False,
      'kind_counts':m.KINDS,'identity_counts':m.GROUPS,
      'blocker_counts':{'FOREIGN_PACKAGING_LIABILITY_NOT_PROVEN':1154,
        'NEW_REVERS_HEADER_NOT_APPROVED':48,'LEGAL_TITLE_NOT_PROVEN':4669,
        'Q00700_PENDING':4669,'HISTORICAL_CANCELLATIONS_NOT_REPLAYED':4669,
        'MIGRATION_ONLY_NOT_IMPORTABLE':4669}}
  for name,key,data in ((m.linked.PRIVATE_HEADERS,'private_header_draft_sha256',headers),
              (m.linked.PRIVATE_ROWS,'private_target_draft_sha256',rows),
              (m.linked.PRIVATE_HOLDS,'private_unknown_document_holds_sha256',holds)):
   data_bytes=''.join(m.canon(item) for item in data).encode()
   (root/name).write_bytes(data_bytes)
   obj[key]=hashlib.sha256(data_bytes).hexdigest()
  (root/'agrix_linked_target_draft_summary.json').write_text(json.dumps(obj),encoding='utf-8')
  return rows,headers,holds
 def test_uuid_stable_not_original(self):
  x=m.candidate_id('P1','1;2')
  self.assertEqual(x,m.candidate_id('P1','1;2'))
  self.assertNotEqual(x,m.candidate_id('P1','1'))
  self.assertNotEqual(x,m.candidate_id('P2','1;2'))
  self.assertTrue(x.startswith('AMB-'))
 def test_big_fixture_partition(self):
  with tempfile.TemporaryDirectory() as t:
   folder=Path(t)/'linked';self.fixture(folder)
   s,mapping,holds,globals=m.build(folder)
   self.assertEqual(len(mapping),4669)
   self.assertEqual(s['candidate_id_collisions'],0)
   self.assertEqual(s['candidate_rows_with_existing_goods_parent'],4621)
   self.assertEqual(s['candidate_rows_with_planned_revers_header'],48)
   self.assertEqual(s['specific_local_gate_counts'],{
    'ISSUANCE_FOREIGN_LIABILITY_POLICY_PENDING':1154,
    'NEW_REVERS_DRAFT_REQUIRES_HEADER_ACCEPTANCE':48})
   self.assertEqual(s['standalone_candidate_rows_without_issuance_or_new_header_gate'],3485)
   self.assertEqual(s['unknown_document_held_events'],6)
   self.assertEqual(s['global_migration_gate_counts'],6)
   self.assertFalse(s['import_approved'])
   self.assertTrue(all(not r['import_approved'] for r in mapping))
   self.assertNotIn('AmbID',mapping[0])
 def test_end_to_end_cli_and_tamper(self):
  with tempfile.TemporaryDirectory() as t:
   root=Path(t);folder=root/'linked';self.fixture(folder)
   out=root/'out';self.assertEqual(m.cli(['--linked',str(folder),'--out',str(out)]),0)
   self.assertTrue((out/m.SUMMARY).is_file())
   self.assertEqual(len((out/m.PRIVATE_MAP).read_text().splitlines()),4669)
   self.assertEqual(m.cli.__name__,'cli')
   with self.assertRaises(AuditError):m.cli(['--linked',str(folder),'--out',str(out)])
   target=folder/m.linked.PRIVATE_ROWS
   with target.open('a',encoding='utf-8') as f:f.write('\n')
   with self.assertRaises(AuditError):m.build(folder)
 def test_unknown_never_becomes_revers(self):
  with tempfile.TemporaryDirectory() as t:
   folder=Path(t)/'linked';rows,heads,holds=self.fixture(folder)
   bad=holds[0].copy();bad['reason']='AUTOMATIC_REVERS'
   with self.assertRaises(AuditError):m.classify(rows,heads,[bad]+holds[1:])
 def test_duplicate_source_leg_or_id_rejected(self):
  with tempfile.TemporaryDirectory() as t:
   folder=Path(t)/'linked';rows,heads,holds=self.fixture(folder)
   q=json.loads(json.dumps(rows));q[1]['source_row_ordinals']=q[0]['source_row_ordinals']
   with self.assertRaises(AuditError):m.classify(q,heads,holds)
 def test_baseline_inflation_rejected(self):
  with tempfile.TemporaryDirectory() as t:
   folder=Path(t)/'linked';self.fixture(folder)
   s=json.loads((folder/'agrix_linked_target_draft_summary.json').read_text());s['source_negative_account_types_allowed']=34
   (folder/'agrix_linked_target_draft_summary.json').write_text(json.dumps(s))
   with self.assertRaises(AuditError):m.verify_linked(folder)

if __name__=='__main__':unittest.main()
