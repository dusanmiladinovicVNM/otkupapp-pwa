import json, tempfile, unittest
from pathlib import Path
from unittest.mock import patch
from legacy_audit import AuditError
import legacy_amb10_ledger_package_v2 as m

class TestMIG063(unittest.TestCase):
 def event(self, case='P1', kind='IZDATA_PRAZNA',fr='Stanica',to='Kooperant',label='OTK-1'):
  return {'physical_case_id':case,'source_row_ordinals':'1;2',
     'canonical_AMB10_import_authorized':'NO','legacy_Datum':'2026-08-10',
     'Kolicina':'50','TipAmbalaze':'Gajba','OdNalogTip':fr,
     'OdNalogID':'ST-1' if fr=='Stanica' else ('K-1' if fr=='Kooperant' else 'V-1'),
     'NaNalogTip':to,'NaNalogID':'ST-1' if to=='Stanica' else ('K-1' if to=='Kooperant' else 'V-1'),
     'legacy_DokumentID':label,'legacy_DokumentTIP':'OM-Izlaz-Koop'}
 def other(self,ev,cat=m.bridge.CATEGORIES[1],kind='IZDATA_PRAZNA'):
  rec={'physical_case_id':ev['physical_case_id'],'source_row_ordinals':ev['source_row_ordinals'],
    'identity_status':cat,'import_approved':False,
    'verified_existing_DokumentTIP':'Otkup' if cat in m.bridge.CATEGORIES[:2] else None,
    'verified_existing_DokumentID':ev['legacy_DokumentID'] if cat in m.bridge.CATEGORIES[:2] else None}
  biz={'physical_case_id':ev['physical_case_id'],'source_row_ordinals':ev['source_row_ordinals'],
       'import_approved':False,'proposed_kind':kind,'business_intent_evidence':'SOURCE_WRITER'}
  return rec,biz
 def test_good_parent_target_shaped_but_not_importable(self):
  ev=self.event();i,b=self.other(ev)
  row,hold=m.target_row(ev,b,i,None)
  self.assertIsNone(hold)
  self.assertEqual(set(row['proposed_target_columns']),set(m.CANONICAL_COLUMNS))
  self.assertEqual(row['proposed_target_columns']['DokumentID'],'OTK-1')
  self.assertIsNone(row['proposed_target_columns']['AmbID'])
  self.assertIsNone(row['proposed_target_columns']['StornoOd'])
  self.assertEqual(row['proposed_target_columns']['Datum'],'2026-08-10')
  self.assertIn('LEGAL_TITLE_NOT_PROVEN',row['import_blockers'])
  self.assertIn('FOREIGN_PACKAGING_LIABILITY_NOT_PROVEN',row['import_blockers'])
  self.assertFalse(row['import_approved'])
 def test_revers_joins_draft_header(self):
  ev=self.event(case='P2',label='3/100826');i,b=self.other(ev,m.bridge.CATEGORIES[2])
  link={'physical_case_id':'P2','source_row_ordinals':'1;2','draft_DokumentTIP':'AmbalazaDokument',
       'draft_DokumentID':'ADK-FAKE','draft_VrstaKretanja':'IZDATA_PRAZNA',
       'title_or_ownership_verified':False,'import_approved':False}
  row,hold=m.target_row(ev,b,i,link)
  self.assertIsNone(hold)
  self.assertEqual(row['proposed_target_columns']['DokumentTIP'],'AmbalazaDokument')
  self.assertEqual(row['proposed_target_columns']['DokumentID'],'ADK-FAKE')
  self.assertIn('NEW_REVERS_HEADER_NOT_APPROVED',row['import_blockers'])
 def test_unknown_is_held_not_invented(self):
  ev=self.event();i,b=self.other(ev,m.bridge.CATEGORIES[3],None)
  x,y=m.target_row(ev,b,i,None)
  self.assertIsNone(x);self.assertEqual(y['reason'],'NO_SOURCE_DOCUMENT_ID_AND_NO_PROVEN_KIND')
  self.assertFalse(y['import_approved'])
 def test_invalid_role_and_approval_fail_closed(self):
  ev=self.event(fr='Stanica',to='Vozac');i,b=self.other(ev)
  with self.assertRaises(AuditError):m.target_row(ev,b,i,None)
  ev=self.event();i,b=self.other(ev);b['import_approved']=True
  with self.assertRaises(AuditError):m.target_row(ev,b,i,None)
  ev=self.event();i,b=self.other(ev,m.bridge.CATEGORIES[2]);
  with self.assertRaises(AuditError):m.target_row(ev,b,i,None)
 def test_inv04_inv10_diagnosis_preserves_cases(self):
  e=self.event();i,b=self.other(e);first,_=m.target_row(e,b,i,None)
  second=json.loads(json.dumps(first));second['physical_case_id']='P2'
  self.assertEqual(m.collisions([first,second]),(1,0,2))
  self.assertIn('AMB_INV_04_COLLISION',first['import_blockers'])
  ev=self.event(kind='POVRAT_PRAZNE',fr='Kooperant',to='Stanica')
  i,b=self.other(ev,kind='POVRAT_PRAZNE');third,_=m.target_row(ev,b,i,None)
  # same pair (unordered) does not cause INV-10 collision
  self.assertEqual(m.collisions([first,third])[1],0)
 def test_existing_output_rejected(self):
  with tempfile.TemporaryDirectory() as d:
   argv=[]
   for key in ('stage','expanded','seven','dryrun','storno-archive','identity','business-intent','revers-plan','out'):
    argv+=['--'+key,d]
   with self.assertRaises(AuditError):m.cli(argv)
 def test_integration_4675_events_and_no_source_leg_loss(self):
  with tempfile.TemporaryDirectory() as temp:
   root=Path(temp);d=root/'revers-plan';d.mkdir()
   paths={k:root/k for k in ('stage','expanded','seven','dryrun','storno-archive','identity','business-intent','revers-plan')}
   for key,p in paths.items():p.mkdir(exist_ok=True)
   summary={'source_sha256':'PINNED','schema_sha256':'SCHEMA',
      'negative_source_account_types_accepted':35,'unexplained_source_balance_residuals':0,
      'historical_cancelled_originals_archived':335,'q00700_held':True}
   nums=iter(n for n in range(1,7481) if n!=700)
   events=[];idents=[];business=[];headers=[];links=[];unknown=[]
   for j in range(4675):
    if j<3485:cat=m.bridge.CATEGORIES[0]
    elif j<4621:cat=m.bridge.CATEGORIES[1]
    elif j<4669:cat=m.bridge.CATEGORIES[2]
    else:cat=m.bridge.CATEGORIES[3]
    case='PHY-'+str(j)
    ords=str(next(nums))+((';'+str(next(nums))) if j<2804 else '')
    fr,to=('Kooperant','Stanica') if j<3485 else ('Stanica','Kooperant')
    if j>=4669: fr,to='Stanica','Vozac'
    ev=self.event(case=case,fr=fr,to=to,label='OTK-'+str(j))
    ev['source_row_ordinals']=ords
    idrow,bizrow=self.other(ev,cat,'AMBALAZA_UZ_ROBU' if j<3485 else ('IZDATA_PRAZNA' if j<4669 else None))
    events.append(ev);idents.append(idrow);business.append(bizrow)
    if cat==m.bridge.CATEGORIES[2]:
     newid='ADK-'+str(j);headers.append({'target_header_draft':{'AmbDokID':newid}})
     links.append({'physical_case_id':case,'source_row_ordinals':ords,
        'draft_DokumentTIP':'AmbalazaDokument','draft_DokumentID':newid,
        'draft_VrstaKretanja':'IZDATA_PRAZNA','title_or_ownership_verified':False,
        'import_approved':False})
    elif cat==m.bridge.CATEGORIES[3]:unknown.append({'physical_case_id':case})
   self.assertEqual(len({int(z) for ev in events for z in ev['source_row_ordinals'].split(';')}),7479)
   with patch.object(m.plan,'build',return_value=(summary,headers,links,unknown)),\
        patch.object(m,'verify_draft',return_value={'private_draft_headers_sha256':'HDR','private_draft_links_sha256':'LINK'}),\
        patch.object(m.chrono,'_verify_inputs',return_value=(None,events,'HASH')),\
        patch.object(m.bridge,'build',return_value=(None,idents)),\
        patch.object(m.intent,'build',return_value=(None,business)):
    s,h,rows,holds=m.build(*[paths[k] for k in ('stage','expanded','seven','dryrun','storno-archive','identity','business-intent','revers-plan')])
    self.assertEqual((len(h),len(rows),len(holds)),(48,4669,6))
    self.assertEqual(s['proposal_rows_with_inv04_or_inv10_hold'],0)
    self.assertEqual(s['source_active_legs_covered'],7479)
    self.assertEqual(s['new_revers_rows_linked'],48)
    self.assertEqual(s['rows_with_existing_parent'],4621)
    self.assertEqual(s['approved_canonical_amb_ids'],0)
    self.assertEqual(s['kind_counts']['IZDATA_PRAZNA'],1184)
    self.assertEqual(s['historical_cancelled_legacy_rows_automatically_replayed'],0)

if __name__=='__main__':unittest.main()
